unit xpgParseChart;

// Copyright (c) 2010,2012 Axolot Data
// Web : http://www.axolot.com/xpg
// Mail: xpg@axolot.com
//
// X   X  PPP    GGG
//  X X   P  P  G
//   X    PPP   G  GG
//  X X   P     G   G
// X   X  P      GGG
//
// File generated with Axolot XPG, Xml Parser Generator.
// Version 0.00.90.
// File created on 2012-10-30 16:03:24

// TODO If an element has required attributs and there also is assigned elements,
//      FAssigneds is repaced with [xaElements] when [xaElements] shall be added
//      to FAssigneds. See _TODO_001_

{$MINENUMSIZE 2}
{$BOOLEVAL OFF}
{$RANGECHECKS OFF}

{.$define XLS_WRITE_EXTLST}

interface

uses Classes, SysUtils, Contnrs, IniFiles, Math,
     xpgPUtils, xpgPLists, xpgPXMLUtils, xpgPXML, xpgParseDrawingCommon,
     xpgPSimpleDOM,
     XLSUtils5;

type TST_RadarStyle =  (strsStandard,strsMarker,strsFilled);
const StrTST_RadarStyle: array[0..2] of AxUCString = ('standard','marker','filled');
type TST_BarDir =  (stbdBar,stbdCol);
const StrTST_BarDir: array[0..1] of AxUCString = ('bar','col');
type TST_SplitType =  (ststAuto,ststCust,ststPercent,ststPos,ststVal);
const StrTST_SplitType: array[0..4] of AxUCString = ('auto','cust','percent','pos','val');
type TST_TickMark =  (sttmCross,sttmIn,sttmNone,sttmOut);
const StrTST_TickMark: array[0..3] of AxUCString = ('cross','in','none','out');
type TST_TimeUnit =  (sttuDays,sttuMonths,sttuYears);
const StrTST_TimeUnit: array[0..2] of AxUCString = ('days','months','years');
type TST_Grouping =  (stgPercentStacked,stgStandard,stgStacked);
const StrTST_Grouping: array[0..2] of AxUCString = ('percentStacked','standard','stacked');
type TST_Orientation =  (stoMaxMin,stoMinMax);
const StrTST_Orientation: array[0..1] of AxUCString = ('maxMin','minMax');
type TST_LblAlgn =  (stlaCtr,stlaL,stlaR);
const StrTST_LblAlgn: array[0..2] of AxUCString = ('ctr','l','r');
type TST_MarkerStyle =  (stmsCircle,stmsDash,stmsDiamond,stmsDot,stmsNone,stmsPicture,stmsPlus,stmsSquare,stmsStar,stmsTriangle,stmsX);
const StrTST_MarkerStyle: array[0..10] of AxUCString = ('circle','dash','diamond','dot','none','picture','plus','square','star','triangle','x');
type TST_Crosses =  (stcAutoZero,stcMax,stcMin);
const StrTST_Crosses: array[0..2] of AxUCString = ('autoZero','max','min');
type TST_PictureFormat =  (stpfStretch,stpfStack,stpfStackScale);
const StrTST_PictureFormat: array[0..2] of AxUCString = ('stretch','stack','stackScale');
type TST_ErrBarType =  (stebtBoth,stebtMinus,stebtPlus);
const StrTST_ErrBarType: array[0..2] of AxUCString = ('both','minus','plus');
type TST_TickLblPos =  (sttlpHigh,sttlpLow,sttlpNextTo,sttlpNone);
const StrTST_TickLblPos: array[0..3] of AxUCString = ('high','low','nextTo','none');
type TST_AxPos =  (stapB,stapL,stapR,stapT);
const StrTST_AxPos: array[0..3] of AxUCString = ('b','l','r','t');
type TST_Shape =  (stsCone,stsConeToMax,stsBox,stsCylinder,stsPyramid,stsPyramidToMax);
const StrTST_Shape: array[0..5] of AxUCString = ('cone','coneToMax','box','cylinder','pyramid','pyramidToMax');
type TST_CrossBetween =  (stcbBetween,stcbMidCat);
const StrTST_CrossBetween: array[0..1] of AxUCString = ('between','midCat');
type TST_DispBlanksAs =  (stdbaSpan,stdbaGap,stdbaZero);
const StrTST_DispBlanksAs: array[0..2] of AxUCString = ('span','gap','zero');
type TST_ErrValType =  (stevtCust,stevtFixedVal,stevtPercentage,stevtStdDev,stevtStdErr);
const StrTST_ErrValType: array[0..4] of AxUCString = ('cust','fixedVal','percentage','stdDev','stdErr');
type TST_LegendPos =  (stlpB,stlpTr,stlpL,stlpR,stlpT);
const StrTST_LegendPos: array[0..4] of AxUCString = ('b','tr','l','r','t');
type TST_ErrDir =  (stedX,stedY);
const StrTST_ErrDir: array[0..1] of AxUCString = ('x','y');
type TST_LayoutMode =  (stlmEdge,stlmFactor);
const StrTST_LayoutMode: array[0..1] of AxUCString = ('edge','factor');
type TST_BarGrouping =  (stbgPercentStacked,stbgClustered,stbgStandard,stbgStacked);
const StrTST_BarGrouping: array[0..3] of AxUCString = ('percentStacked','clustered','standard','stacked');
type TST_DLblPos =  (stdlpBestFit,stdlpB,stdlpCtr,stdlpInBase,stdlpInEnd,stdlpL,stdlpOutEnd,stdlpR,stdlpT);
const StrTST_DLblPos: array[0..8] of AxUCString = ('bestFit','b','ctr','inBase','inEnd','l','outEnd','r','t');
type TST_SizeRepresents =  (stsrArea,stsrW);
const StrTST_SizeRepresents: array[0..1] of AxUCString = ('area','w');
type TST_PageSetupOrientation =  (stpsoDefault,stpsoPortrait,stpsoLandscape);
const StrTST_PageSetupOrientation: array[0..2] of AxUCString = ('default','portrait','landscape');
type TST_OfPieType =  (stoptPie,stoptBar);
const StrTST_OfPieType: array[0..1] of AxUCString = ('pie','bar');
type TST_ScatterStyle =  (stssNone,stssLine,stssLineMarker,stssMarker,stssSmooth,stssSmoothMarker);
const StrTST_ScatterStyle: array[0..5] of AxUCString = ('none','line','lineMarker','marker','smooth','smoothMarker');
type TST_LayoutTarget =  (stltInner,stltOuter);
const StrTST_LayoutTarget: array[0..1] of AxUCString = ('inner','outer');
type TST_BuiltInUnit =  (stbiuHundreds,stbiuThousands,stbiuTenThousands,stbiuHundredThousands,stbiuMillions,stbiuTenMillions,stbiuHundredMillions,stbiuBillions,stbiuTrillions);
const StrTST_BuiltInUnit: array[0..8] of AxUCString = ('hundreds','thousands','tenThousands','hundredThousands','millions','tenMillions','hundredMillions','billions','trillions');
type TST_TrendlineType =  (stttExp,stttLinear,stttLog,stttMovingAvg,stttPoly,stttPower);
const StrTST_TrendlineType: array[0..5] of AxUCString = ('exp','linear','log','movingAvg','poly','power');

// Not the same as a drawing T_CTShape
type TCT_Shape_Chart = class(TXPGBase)
protected
     FVal: TST_Shape;
public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Val: TST_Shape read FVal write FVal;
     end;


     TCT_Extension = class(TXPGBase)
protected
     FUri: AxUCString;
     FXmlns: AxUCString;
     FXmlnsVal: AxUCString;
     FAnyElements: TXPGAnyElements;
public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Extension);
     procedure CopyTo(AItem: TCT_Extension);

     property Uri: AxUCString read FUri write FUri;
     property AnyElements: TXPGAnyElements read FAnyElements;
     end;

     TCT_ExtensionXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_Extension;
public
     function  Add: TCT_Extension;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_ExtensionXpgList);
     procedure CopyTo(AItem: TCT_ExtensionXpgList);
     property Items[Index: integer]: TCT_Extension read GetItems; default;
     end;

     TCT_UnsignedInt = class(TXPGBase)
protected
     FVal: integer;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_UnsignedInt);
     procedure CopyTo(AItem: TCT_UnsignedInt);

     property Val: integer read FVal write FVal;
     end;

     TCT_UnsignedIntXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_UnsignedInt;
public
     function  Add: TCT_UnsignedInt;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_UnsignedIntXpgList);
     procedure CopyTo(AItem: TCT_UnsignedIntXpgList);
     property Items[Index: integer]: TCT_UnsignedInt read GetItems; default;
     end;

     TCT_StrVal = class(TXPGBase)
protected
     FIdx: integer;
     FV: AxUCString;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_StrVal);
     procedure CopyTo(AItem: TCT_StrVal);

     property Idx: integer read FIdx write FIdx;
     property V: AxUCString read FV write FV;
     end;

     TCT_StrValXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_StrVal;
public
     function  Add: TCT_StrVal;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_StrValXpgList);
     procedure CopyTo(AItem: TCT_StrValXpgList);
     property Items[Index: integer]: TCT_StrVal read GetItems; default;
     end;

     TCT_ExtensionList = class(TXPGBase)
protected
     FExtXpgList: TCT_ExtensionXpgList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_ExtensionList);
     procedure CopyTo(AItem: TCT_ExtensionList);
     function  Create_ExtXpgList: TCT_ExtensionXpgList;

     property ExtXpgList: TCT_ExtensionXpgList read FExtXpgList;
     end;

     TCT_ExtensionListXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_ExtensionList;
public
     function  Add: TCT_ExtensionList;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_ExtensionListXpgList);
     procedure CopyTo(AItem: TCT_ExtensionListXpgList);
     property Items[Index: integer]: TCT_ExtensionList read GetItems; default;
     end;

     TCT_LayoutTarget = class(TXPGBase)
protected
     FVal: TST_LayoutTarget;
public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_LayoutTarget);
     procedure CopyTo(AItem: TCT_LayoutTarget);

     property Val: TST_LayoutTarget read FVal write FVal;
     end;

     TCT_LayoutMode = class(TXPGBase)
protected
     FVal: TST_LayoutMode;
public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_LayoutMode);
     procedure CopyTo(AItem: TCT_LayoutMode);

     property Val: TST_LayoutMode read FVal write FVal;
     end;

     TCT_Double = class(TXPGBase)
protected
     FVal: double;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Double);
     procedure CopyTo(AItem: TCT_Double);

     property Val: double read FVal write FVal;
     end;

     TCT_StrData = class(TXPGBase)
protected
     FPtCount: TCT_UnsignedInt;
     FPtXpgList: TCT_StrValXpgList;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_StrData);
     procedure CopyTo(AItem: TCT_StrData);
     function  Create_PtCount: TCT_UnsignedInt;
     function  Create_PtXpgList: TCT_StrValXpgList;
     function  Create_ExtLst: TCT_ExtensionList;

     property PtCount: TCT_UnsignedInt read FPtCount;
     property PtXpgList: TCT_StrValXpgList read FPtXpgList;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_ManualLayout = class(TXPGBase)
protected
     FLayoutTarget: TCT_LayoutTarget;
     FXMode: TCT_LayoutMode;
     FYMode: TCT_LayoutMode;
     FWMode: TCT_LayoutMode;
     FHMode: TCT_LayoutMode;
     FX: TCT_Double;
     FY: TCT_Double;
     FW: TCT_Double;
     FH: TCT_Double;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_ManualLayout);
     procedure CopyTo(AItem: TCT_ManualLayout);
     function  Create_LayoutTarget: TCT_LayoutTarget;
     function  Create_XMode: TCT_LayoutMode;
     function  Create_YMode: TCT_LayoutMode;
     function  Create_WMode: TCT_LayoutMode;
     function  Create_HMode: TCT_LayoutMode;
     function  Create_X: TCT_Double;
     function  Create_Y: TCT_Double;
     function  Create_W: TCT_Double;
     function  Create_H: TCT_Double;
     function  Create_ExtLst: TCT_ExtensionList;

     property LayoutTarget: TCT_LayoutTarget read FLayoutTarget;
     property XMode: TCT_LayoutMode read FXMode;
     property YMode: TCT_LayoutMode read FYMode;
     property WMode: TCT_LayoutMode read FWMode;
     property HMode: TCT_LayoutMode read FHMode;
     property X: TCT_Double read FX;
     property Y: TCT_Double read FY;
     property W: TCT_Double read FW;
     property H: TCT_Double read FH;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_StrRef = class(TXPGBase)
protected
     FF: AxUCString;
     FStrCache: TCT_StrData;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_StrRef);
     procedure CopyTo(AItem: TCT_StrRef);
     function  Create_StrCache: TCT_StrData;
     function  Create_ExtLst: TCT_ExtensionList;

     property F: AxUCString read FF write FF;
     property StrCache: TCT_StrData read FStrCache;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_NumFmt = class(TXPGBase)
protected
     FFormatCode: AxUCString;
     FSourceLinked: boolean;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_NumFmt);
     procedure CopyTo(AItem: TCT_NumFmt);

     property FormatCode: AxUCString read FFormatCode write FFormatCode;
     property SourceLinked: boolean read FSourceLinked write FSourceLinked;
     end;

     TCT_DLblPos = class(TXPGBase)
protected
     FVal: TST_DLblPos;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_DLblPos);
     procedure CopyTo(AItem: TCT_DLblPos);

     property Val: TST_DLblPos read FVal write FVal;
     end;

     TCT_Boolean = class(TXPGBase)
protected
     FVal: boolean;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Boolean);
     procedure CopyTo(AItem: TCT_Boolean);

     property Val: boolean read FVal write FVal;
     end;

     TCT_NumVal = class(TXPGBase)
protected
     FIdx: integer;
     FFormatCode: AxUCString;
     FV: AxUCString;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_NumVal);
     procedure CopyTo(AItem: TCT_NumVal);

     property Idx: integer read FIdx write FIdx;
     property FormatCode: AxUCString read FFormatCode write FFormatCode;
     property V: AxUCString read FV write FV;
     end;

     TCT_NumValXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_NumVal;
public
     function  Add: TCT_NumVal;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_NumValXpgList);
     procedure CopyTo(AItem: TCT_NumValXpgList);
     property Items[Index: integer]: TCT_NumVal read GetItems; default;
     end;

     TCT_Layout = class(TXPGBase)
protected
     FX: double;
     FY: double;
     FW: double;
     FH: double;

     FManualLayout: TCT_ManualLayout;
     FExtLst: TCT_ExtensionList;
public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Prepare(const AX,AY,AW,AH: double);
     procedure Assign(AItem: TCT_Layout);
     procedure CopyTo(AItem: TCT_Layout);
     function  Create_ManualLayout: TCT_ManualLayout;
     function  Create_ExtLst: TCT_ExtensionList;

     property ManualLayout: TCT_ManualLayout read FManualLayout;
     property ExtLst: TCT_ExtensionList read FExtLst;

     property X: double read FX write FX;
     property Y: double read FY write FY;
     property W: double read FW write FW;
     property H: double read FH write FH;
     end;

     TCT_Tx = class(TXPGBase)
protected
     FStrRef: TCT_StrRef;
     FRich: TCT_TextBody;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Tx);
     procedure CopyTo(AItem: TCT_Tx);
     function  Create_StrRef: TCT_StrRef;
     function  Create_Rich: TCT_TextBody;

     property StrRef: TCT_StrRef read FStrRef;
     property Rich: TCT_TextBody read FRich;
     end;

     TEG_DLblShared = class(TXPGBase)
protected
     FNumFmt: TCT_NumFmt;
     FSpPr: TCT_ShapeProperties;
     FTxPr: TCT_TextBody;
     FDLblPos: TCT_DLblPos;
     FShowLegendKey: TCT_Boolean;
     FShowVal: TCT_Boolean;
     FShowCatName: TCT_Boolean;
     FShowSerName: TCT_Boolean;
     FShowPercent: TCT_Boolean;
     FShowBubbleSize: TCT_Boolean;
     FSeparator: AxUCString;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TEG_DLblShared);
     procedure CopyTo(AItem: TEG_DLblShared);
     function  Create_NumFmt: TCT_NumFmt;
     function  Create_SpPr: TCT_ShapeProperties;
     function  Create_TxPr: TCT_TextBody;
     function  Create_DLblPos: TCT_DLblPos;
     function  Create_ShowLegendKey: TCT_Boolean;
     function  Create_ShowVal: TCT_Boolean;
     function  Create_ShowCatName: TCT_Boolean;
     function  Create_ShowSerName: TCT_Boolean;
     function  Create_ShowPercent: TCT_Boolean;
     function  Create_ShowBubbleSize: TCT_Boolean;

     property NumFmt: TCT_NumFmt read FNumFmt;
     property SpPr: TCT_ShapeProperties read FSpPr;
     property TxPr: TCT_TextBody read FTxPr;
     property DLblPos: TCT_DLblPos read FDLblPos;
     property ShowLegendKey: TCT_Boolean read FShowLegendKey;
     property ShowVal: TCT_Boolean read FShowVal;
     property ShowCatName: TCT_Boolean read FShowCatName;
     property ShowSerName: TCT_Boolean read FShowSerName;
     property ShowPercent: TCT_Boolean read FShowPercent;
     property ShowBubbleSize: TCT_Boolean read FShowBubbleSize;
     property Separator: AxUCString read FSeparator write FSeparator;
     end;

     TCT_NumData = class(TXPGBase)
protected
     FFormatCode: AxUCString;
     FPtCount: TCT_UnsignedInt;
     FPtXpgList: TCT_NumValXpgList;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_NumData);
     procedure CopyTo(AItem: TCT_NumData);
     function  Create_PtCount: TCT_UnsignedInt;
     function  Create_PtXpgList: TCT_NumValXpgList;
     function  Create_ExtLst: TCT_ExtensionList;

     property FormatCode: AxUCString read FFormatCode write FFormatCode;
     property PtCount: TCT_UnsignedInt read FPtCount;
     property PtXpgList: TCT_NumValXpgList read FPtXpgList;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_Lvl = class(TXPGBase)
protected
     FPtXpgList: TCT_StrValXpgList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Lvl);
     procedure CopyTo(AItem: TCT_Lvl);
     function  Create_PtXpgList: TCT_StrValXpgList;

     property PtXpgList: TCT_StrValXpgList read FPtXpgList;
     end;

     TCT_LvlXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_Lvl;
public
     function  Add: TCT_Lvl;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_LvlXpgList);
     procedure CopyTo(AItem: TCT_LvlXpgList);
     property Items[Index: integer]: TCT_Lvl read GetItems; default;
     end;

     TCT_MarkerStyle = class(TXPGBase)
protected
     FVal: TST_MarkerStyle;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_MarkerStyle);
     procedure CopyTo(AItem: TCT_MarkerStyle);

     property Val: TST_MarkerStyle read FVal write FVal;
     end;

     TCT_MarkerSize = class(TXPGBase)
protected
     FVal: integer;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_MarkerSize);
     procedure CopyTo(AItem: TCT_MarkerSize);

     property Val: integer read FVal write FVal;
     end;

     TCT_PictureFormat = class(TXPGBase)
protected
     FVal: TST_PictureFormat;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_PictureFormat);
     procedure CopyTo(AItem: TCT_PictureFormat);

     property Val: TST_PictureFormat read FVal write FVal;
     end;

     TCT_PictureStackUnit = class(TXPGBase)
protected
     FVal: double;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_PictureStackUnit);
     procedure CopyTo(AItem: TCT_PictureStackUnit);

     property Val: double read FVal write FVal;
     end;

     TGroup_DLbl = class(TXPGBase)
protected
     FLayout: TCT_Layout;
     FTx: TCT_Tx;
     FEG_DLblShared: TEG_DLblShared;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TGroup_DLbl);
     procedure CopyTo(AItem: TGroup_DLbl);
     function  Create_Layout: TCT_Layout;
     function  Create_Tx: TCT_Tx;

     property Layout: TCT_Layout read FLayout;
     property Tx: TCT_Tx read FTx;
     property EG_DLblShared: TEG_DLblShared read FEG_DLblShared;
     end;

     TCT_ChartLines = class(TXPGBase)
protected
     FSpPr: TCT_ShapeProperties;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_ChartLines);
     procedure CopyTo(AItem: TCT_ChartLines);
     function  Create_SpPr: TCT_ShapeProperties;

     property SpPr: TCT_ShapeProperties read FSpPr;
     end;

     TCT_ChartLinesXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_ChartLines;
public
     function  Add: TCT_ChartLines;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_ChartLinesXpgList);
     procedure CopyTo(AItem: TCT_ChartLinesXpgList);
     property Items[Index: integer]: TCT_ChartLines read GetItems; default;
     end;

     TCT_NumRef = class(TXPGBase)
protected
     FF: AxUCString;
     FNumCache: TCT_NumData;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_NumRef);
     procedure CopyTo(AItem: TCT_NumRef);
     function  Create_NumCache: TCT_NumData;
     function  Create_ExtLst: TCT_ExtensionList;

     property F: AxUCString read FF write FF;
     property NumCache: TCT_NumData read FNumCache;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_MultiLvlStrData = class(TXPGBase)
protected
     FPtCount: TCT_UnsignedInt;
     FLvlXpgList: TCT_LvlXpgList;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_MultiLvlStrData);
     procedure CopyTo(AItem: TCT_MultiLvlStrData);
     function  Create_PtCount: TCT_UnsignedInt;
     function  Create_LvlXpgList: TCT_LvlXpgList;
     function  Create_ExtLst: TCT_ExtensionList;

     property PtCount: TCT_UnsignedInt read FPtCount;
     property LvlXpgList: TCT_LvlXpgList read FLvlXpgList;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_SerTx = class(TXPGBase)
protected
     FStrRef: TCT_StrRef;
     FV: AxUCString;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_SerTx);
     procedure CopyTo(AItem: TCT_SerTx);
     function  Create_StrRef: TCT_StrRef;

     property StrRef: TCT_StrRef read FStrRef;
     property V: AxUCString read FV write FV;
     end;

     TCT_PictureOptions = class(TXPGBase)
protected
     FApplyToFront: TCT_Boolean;
     FApplyToSides: TCT_Boolean;
     FApplyToEnd: TCT_Boolean;
     FPictureFormat: TCT_PictureFormat;
     FPictureStackUnit: TCT_PictureStackUnit;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_PictureOptions);
     procedure CopyTo(AItem: TCT_PictureOptions);
     function  Create_ApplyToFront: TCT_Boolean;
     function  Create_ApplyToSides: TCT_Boolean;
     function  Create_ApplyToEnd: TCT_Boolean;
     function  Create_PictureFormat: TCT_PictureFormat;
     function  Create_PictureStackUnit: TCT_PictureStackUnit;

     property ApplyToFront: TCT_Boolean read FApplyToFront;
     property ApplyToSides: TCT_Boolean read FApplyToSides;
     property ApplyToEnd: TCT_Boolean read FApplyToEnd;
     property PictureFormat: TCT_PictureFormat read FPictureFormat;
     property PictureStackUnit: TCT_PictureStackUnit read FPictureStackUnit;
     end;

     TCT_DLbl = class(TXPGBase)
protected
     FIdx: TCT_UnsignedInt;
     FDelete: TCT_Boolean;
     FGroup_DLbl: TGroup_DLbl;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_DLbl);
     procedure CopyTo(AItem: TCT_DLbl);
     function  Create_Idx: TCT_UnsignedInt;
     function  Create_Delete: TCT_Boolean;
     function  Create_ExtLst: TCT_ExtensionList;

     property Idx: TCT_UnsignedInt read FIdx;
     property Delete: TCT_Boolean read FDelete;
     property Group_DLbl: TGroup_DLbl read FGroup_DLbl;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_DLblXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_DLbl;
public
     function  Add: TCT_DLbl;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_DLblXpgList);
     procedure CopyTo(AItem: TCT_DLblXpgList);
     property Items[Index: integer]: TCT_DLbl read GetItems; default;
     end;

     TGroup_DLbls = class(TXPGBase)
protected
     FEG_DLblShared: TEG_DLblShared;
     FShowLeaderLines: TCT_Boolean;
     FLeaderLines: TCT_ChartLines;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TGroup_DLbls);
     procedure CopyTo(AItem: TGroup_DLbls);
     function  Create_ShowLeaderLines: TCT_Boolean;
     function  Create_LeaderLines: TCT_ChartLines;

     property EG_DLblShared: TEG_DLblShared read FEG_DLblShared;
     property ShowLeaderLines: TCT_Boolean read FShowLeaderLines;
     property LeaderLines: TCT_ChartLines read FLeaderLines;
     end;

     TCT_TrendlineType = class(TXPGBase)
protected
     FVal: TST_TrendlineType;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_TrendlineType);
     procedure CopyTo(AItem: TCT_TrendlineType);

     property Val: TST_TrendlineType read FVal write FVal;
     end;

     TCT_Order = class(TXPGBase)
protected
     FVal: integer;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Order);
     procedure CopyTo(AItem: TCT_Order);

     property Val: integer read FVal write FVal;
     end;

     TCT_Period = class(TXPGBase)
protected
     FVal: integer;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Period);
     procedure CopyTo(AItem: TCT_Period);

     property Val: integer read FVal write FVal;
     end;

     TCT_TrendlineLbl = class(TXPGBase)
protected
     FLayout: TCT_Layout;
     FTx: TCT_Tx;
     FNumFmt: TCT_NumFmt;
     FSpPr: TCT_ShapeProperties;
     FTxPr: TCT_TextBody;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_TrendlineLbl);
     procedure CopyTo(AItem: TCT_TrendlineLbl);
     function  Create_Layout: TCT_Layout;
     function  Create_Tx: TCT_Tx;
     function  Create_NumFmt: TCT_NumFmt;
     function  Create_SpPr: TCT_ShapeProperties;
     function  Create_TxPr: TCT_TextBody;
     function  Create_ExtLst: TCT_ExtensionList;

     property Layout: TCT_Layout read FLayout;
     property Tx: TCT_Tx read FTx;
     property NumFmt: TCT_NumFmt read FNumFmt;
     property SpPr: TCT_ShapeProperties read FSpPr;
     property TxPr: TCT_TextBody read FTxPr;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_ErrDir = class(TXPGBase)
protected
     FVal: TST_ErrDir;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_ErrDir);
     procedure CopyTo(AItem: TCT_ErrDir);

     property Val: TST_ErrDir read FVal write FVal;
     end;

     TCT_ErrBarType = class(TXPGBase)
protected
     FVal: TST_ErrBarType;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_ErrBarType);
     procedure CopyTo(AItem: TCT_ErrBarType);

     property Val: TST_ErrBarType read FVal write FVal;
     end;

     TCT_ErrValType = class(TXPGBase)
protected
     FVal: TST_ErrValType;
public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Val: TST_ErrValType read FVal write FVal;
     end;

     TCT_NumDataSource = class(TXPGBase)
protected
     FNumRef: TCT_NumRef;
     FNumLit: TCT_NumData;
public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_NumDataSource);
     procedure CopyTo(AItem: TCT_NumDataSource);
     function  Create_NumRef: TCT_NumRef;
     function  Create_NumLit: TCT_NumData;

     property NumRef: TCT_NumRef read FNumRef;
     property NumLit: TCT_NumData read FNumLit;
     end;

     TCT_MultiLvlStrRef = class(TXPGBase)
protected
     FF: AxUCString;
     FMultiLvlStrCache: TCT_MultiLvlStrData;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_MultiLvlStrRef);
     procedure CopyTo(AItem: TCT_MultiLvlStrRef);
     function  Create_MultiLvlStrCache: TCT_MultiLvlStrData;
     function  Create_ExtLst: TCT_ExtensionList;

     property F: AxUCString read FF write FF;
     property MultiLvlStrCache: TCT_MultiLvlStrData read FMultiLvlStrCache;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TEG_SerShared = class(TXPGBase)
protected
     FIdx: TCT_UnsignedInt;
     FOrder: TCT_UnsignedInt;
     FTx: TCT_SerTx;
     FSpPr: TCT_ShapeProperties;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TEG_SerShared);
     procedure CopyTo(AItem: TEG_SerShared);
     function  Create_Idx: TCT_UnsignedInt;
     function  Create_Order: TCT_UnsignedInt;
     function  Create_Tx: TCT_SerTx;
     function  Create_SpPr: TCT_ShapeProperties;

     property Idx: TCT_UnsignedInt read FIdx;
     property Order: TCT_UnsignedInt read FOrder;
     property Tx: TCT_SerTx read FTx;
     property SpPr: TCT_ShapeProperties read FSpPr;
     end;

     TCT_Marker = class(TXPGBase)
protected
     FSymbol: TCT_MarkerStyle;
     FSize: TCT_MarkerSize;
     FSpPr: TCT_ShapeProperties;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Marker);
     procedure CopyTo(AItem: TCT_Marker);
     function  Create_Symbol: TCT_MarkerStyle;
     function  Create_Size: TCT_MarkerSize;
     function  Create_SpPr: TCT_ShapeProperties;
     function  Create_ExtLst: TCT_ExtensionList;

     property Symbol: TCT_MarkerStyle read FSymbol;
     property Size: TCT_MarkerSize read FSize;
     property SpPr: TCT_ShapeProperties read FSpPr;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_DPt = class(TXPGBase)
protected
     FIdx: TCT_UnsignedInt;
     FInvertIfNegative: TCT_Boolean;
     FMarker: TCT_Marker;
     FBubble3D: TCT_Boolean;
     FExplosion: TCT_UnsignedInt;
     FSpPr: TCT_ShapeProperties;
     FPictureOptions: TCT_PictureOptions;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_DPt);
     procedure CopyTo(AItem: TCT_DPt);
     function  Create_Idx: TCT_UnsignedInt;
     function  Create_InvertIfNegative: TCT_Boolean;
     function  Create_Marker: TCT_Marker;
     function  Create_Bubble3D: TCT_Boolean;
     function  Create_Explosion: TCT_UnsignedInt;
     function  Create_SpPr: TCT_ShapeProperties;
     function  Create_PictureOptions: TCT_PictureOptions;
     function  Create_ExtLst: TCT_ExtensionList;

     property Idx: TCT_UnsignedInt read FIdx;
     property InvertIfNegative: TCT_Boolean read FInvertIfNegative;
     property Marker: TCT_Marker read FMarker;
     property Bubble3D: TCT_Boolean read FBubble3D;
     property Explosion: TCT_UnsignedInt read FExplosion;
     property SpPr: TCT_ShapeProperties read FSpPr;
     property PictureOptions: TCT_PictureOptions read FPictureOptions;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_DPtXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_DPt;
public
     function  Add: TCT_DPt;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_DPtXpgList);
     procedure CopyTo(AItem: TCT_DPtXpgList);
     property Items[Index: integer]: TCT_DPt read GetItems; default;
     end;

     TCT_DLbls = class(TXPGBase)
protected
     FDLblXpgList: TCT_DLblXpgList;
     FDelete: TCT_Boolean;
     FGroup_DLbls: TGroup_DLbls;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_DLbls);
     procedure CopyTo(AItem: TCT_DLbls);
     function  Create_DLblXpgList: TCT_DLblXpgList;
     function  Create_Delete: TCT_Boolean;
     function  Create_ExtLst: TCT_ExtensionList;

     property DLblXpgList: TCT_DLblXpgList read FDLblXpgList;
     property Delete: TCT_Boolean read FDelete;
     property Group_DLbls: TGroup_DLbls read FGroup_DLbls;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_Trendline = class(TXPGBase)
protected
     FName: AxUCString;
     FSpPr: TCT_ShapeProperties;
     FTrendlineType: TCT_TrendlineType;
     FOrder: TCT_Order;
     FPeriod: TCT_Period;
     FForward: TCT_Double;
     FBackward: TCT_Double;
     FIntercept: TCT_Double;
     FDispRSqr: TCT_Boolean;
     FDispEq: TCT_Boolean;
     FTrendlineLbl: TCT_TrendlineLbl;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Trendline);
     procedure CopyTo(AItem: TCT_Trendline);
     function  Create_SpPr: TCT_ShapeProperties;
     function  Create_TrendlineType: TCT_TrendlineType;
     function  Create_Order: TCT_Order;
     function  Create_Period: TCT_Period;
     function  Create_Forward: TCT_Double;
     function  Create_Backward: TCT_Double;
     function  Create_Intercept: TCT_Double;
     function  Create_DispRSqr: TCT_Boolean;
     function  Create_DispEq: TCT_Boolean;
     function  Create_TrendlineLbl: TCT_TrendlineLbl;
     function  Create_ExtLst: TCT_ExtensionList;

     property Name: AxUCString read FName write FName;
     property SpPr: TCT_ShapeProperties read FSpPr;
     property TrendlineType: TCT_TrendlineType read FTrendlineType;
     property Order: TCT_Order read FOrder;
     property Period: TCT_Period read FPeriod;
     property Forward: TCT_Double read FForward;
     property Backward: TCT_Double read FBackward;
     property Intercept: TCT_Double read FIntercept;
     property DispRSqr: TCT_Boolean read FDispRSqr;
     property DispEq: TCT_Boolean read FDispEq;
     property TrendlineLbl: TCT_TrendlineLbl read FTrendlineLbl;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_TrendlineXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_Trendline;
public
     function  Add: TCT_Trendline;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_TrendlineXpgList);
     procedure CopyTo(AItem: TCT_TrendlineXpgList);
     property Items[Index: integer]: TCT_Trendline read GetItems; default;
     end;

     TCT_ErrBars = class(TXPGBase)
protected
     FErrDir: TCT_ErrDir;
     FErrBarType: TCT_ErrBarType;
     FErrValType: TCT_ErrValType;
     FNoEndCap: TCT_Boolean;
     FPlus: TCT_NumDataSource;
     FMinus: TCT_NumDataSource;
     FVal: TCT_Double;
     FSpPr: TCT_ShapeProperties;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_ErrBars);
     procedure CopyTo(AItem: TCT_ErrBars);
     function  Create_ErrDir: TCT_ErrDir;
     function  Create_ErrBarType: TCT_ErrBarType;
     function  Create_ErrValType: TCT_ErrValType;
     function  Create_NoEndCap: TCT_Boolean;
     function  Create_Plus: TCT_NumDataSource;
     function  Create_Minus: TCT_NumDataSource;
     function  Create_Val: TCT_Double;
     function  Create_SpPr: TCT_ShapeProperties;
     function  Create_ExtLst: TCT_ExtensionList;

     property ErrDir: TCT_ErrDir read FErrDir;
     property ErrBarType: TCT_ErrBarType read FErrBarType;
     property ErrValType: TCT_ErrValType read FErrValType;
     property NoEndCap: TCT_Boolean read FNoEndCap;
     property Plus: TCT_NumDataSource read FPlus;
     property Minus: TCT_NumDataSource read FMinus;
     property Val: TCT_Double read FVal;
     property SpPr: TCT_ShapeProperties read FSpPr;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_ErrBarsXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_ErrBars;
public
     function  Add: TCT_ErrBars;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_ErrBarsXpgList);
     procedure CopyTo(AItem: TCT_ErrBarsXpgList);
     property Items[Index: integer]: TCT_ErrBars read GetItems; default;
     end;

     TCT_AxDataSource = class(TXPGBase)
protected
     FMultiLvlStrRef: TCT_MultiLvlStrRef;
     FNumRef: TCT_NumRef;
     FNumLit: TCT_NumData;
     FStrRef: TCT_StrRef;
     FStrLit: TCT_StrData;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_AxDataSource);
     procedure CopyTo(AItem: TCT_AxDataSource);
     function  Create_MultiLvlStrRef: TCT_MultiLvlStrRef;
     function  Create_NumRef: TCT_NumRef;
     function  Create_NumLit: TCT_NumData;
     function  Create_StrRef: TCT_StrRef;
     function  Create_StrLit: TCT_StrData;

     property MultiLvlStrRef: TCT_MultiLvlStrRef read FMultiLvlStrRef;
     property NumRef: TCT_NumRef read FNumRef;
     property NumLit: TCT_NumData read FNumLit;
     property StrRef: TCT_StrRef read FStrRef;
     property StrLit: TCT_StrData read FStrLit;
     end;

     TCT_BandFmt = class(TXPGBase)
protected
     FIdx: TCT_UnsignedInt;
     FSpPr: TCT_ShapeProperties;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_BandFmt);
     procedure CopyTo(AItem: TCT_BandFmt);
     function  Create_Idx: TCT_UnsignedInt;
     function  Create_SpPr: TCT_ShapeProperties;

     property Idx: TCT_UnsignedInt read FIdx;
     property SpPr: TCT_ShapeProperties read FSpPr;
     end;

     TCT_BandFmtXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_BandFmt;
public
     function  Add: TCT_BandFmt;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_BandFmtXpgList);
     procedure CopyTo(AItem: TCT_BandFmtXpgList);
     property Items[Index: integer]: TCT_BandFmt read GetItems; default;
     end;

     TCT_LogBase = class(TXPGBase)
protected
     FVal: double;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_LogBase);
     procedure CopyTo(AItem: TCT_LogBase);

     property Val: double read FVal write FVal;
     end;

     TCT_Orientation = class(TXPGBase)
protected
     FVal: TST_Orientation;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Orientation);
     procedure CopyTo(AItem: TCT_Orientation);

     property Val: TST_Orientation read FVal write FVal;
     end;

     TCT_Grouping = class(TXPGBase)
protected
     FVal: TST_Grouping;
public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Grouping);
     procedure CopyTo(AItem: TCT_Grouping);

     property Val: TST_Grouping read FVal write FVal;
     end;

     TCT_AreaSer = class(TXPGBase)
protected
     FEG_SerShared: TEG_SerShared;
     FPictureOptions: TCT_PictureOptions;
     FDPtXpgList: TCT_DPtXpgList;
     FDLbls: TCT_DLbls;
     FTrendlineXpgList: TCT_TrendlineXpgList;
     FErrBarsXpgList: TCT_ErrBarsXpgList;
     FCat: TCT_AxDataSource;
     FVal: TCT_NumDataSource;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_AreaSer);
     procedure CopyTo(AItem: TCT_AreaSer);
     function  Create_PictureOptions: TCT_PictureOptions;
     function  Create_DPtXpgList: TCT_DPtXpgList;
     function  Create_DLbls: TCT_DLbls;
     function  Create_TrendlineXpgList: TCT_TrendlineXpgList;
     function  Create_ErrBarsXpgList: TCT_ErrBarsXpgList;
     function  Create_Cat: TCT_AxDataSource;
     function  Create_Val: TCT_NumDataSource;
     function  Create_ExtLst: TCT_ExtensionList;

     property EG_SerShared: TEG_SerShared read FEG_SerShared;
     property PictureOptions: TCT_PictureOptions read FPictureOptions;
     property DPtXpgList: TCT_DPtXpgList read FDPtXpgList;
     property DLbls: TCT_DLbls read FDLbls;
     property TrendlineXpgList: TCT_TrendlineXpgList read FTrendlineXpgList;
     property ErrBarsXpgList: TCT_ErrBarsXpgList read FErrBarsXpgList;
     property Cat: TCT_AxDataSource read FCat;
     property Val: TCT_NumDataSource read FVal;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_AreaSerXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_AreaSer;
public
     function  Add: TCT_AreaSer;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_AreaSerXpgList);
     procedure CopyTo(AItem: TCT_AreaSerXpgList);
     property Items[Index: integer]: TCT_AreaSer read GetItems; default;
     end;

     TCT_LineSer = class(TXPGBase)
protected
     FEG_SerShared: TEG_SerShared;
     FMarker: TCT_Marker;
     FDPtXpgList: TCT_DPtXpgList;
     FDLbls: TCT_DLbls;
     FTrendlineXpgList: TCT_TrendlineXpgList;
     FErrBars: TCT_ErrBars;
     FCat: TCT_AxDataSource;
     FVal: TCT_NumDataSource;
     FSmooth: TCT_Boolean;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_LineSer);
     procedure CopyTo(AItem: TCT_LineSer);
     function  Create_Marker: TCT_Marker;
     function  Create_DPtXpgList: TCT_DPtXpgList;
     function  Create_DLbls: TCT_DLbls;
     function  Create_TrendlineXpgList: TCT_TrendlineXpgList;
     function  Create_ErrBars: TCT_ErrBars;
     function  Create_Cat: TCT_AxDataSource;
     function  Create_Val: TCT_NumDataSource;
     function  Create_Smooth: TCT_Boolean;
     function  Create_ExtLst: TCT_ExtensionList;

     property EG_SerShared: TEG_SerShared read FEG_SerShared;
     property Marker: TCT_Marker read FMarker;
     property DPtXpgList: TCT_DPtXpgList read FDPtXpgList;
     property DLbls: TCT_DLbls read FDLbls;
     property TrendlineXpgList: TCT_TrendlineXpgList read FTrendlineXpgList;
     property ErrBars: TCT_ErrBars read FErrBars;
     property Cat: TCT_AxDataSource read FCat;
     property Val: TCT_NumDataSource read FVal;
     property Smooth: TCT_Boolean read FSmooth;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_LineSerXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_LineSer;
public
     function  Add: TCT_LineSer;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_LineSerXpgList);
     procedure CopyTo(AItem: TCT_LineSerXpgList);
     property Items[Index: integer]: TCT_LineSer read GetItems; default;
     end;

     TCT_GapAmount = class(TXPGBase)
protected
     FVal: integer;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_GapAmount);
     procedure CopyTo(AItem: TCT_GapAmount);

     property Val: integer read FVal write FVal;
     end;

     TCT_UpDownBar = class(TXPGBase)
protected
     FSpPr: TCT_ShapeProperties;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_UpDownBar);
     procedure CopyTo(AItem: TCT_UpDownBar);
     function  Create_SpPr: TCT_ShapeProperties;

     property SpPr: TCT_ShapeProperties read FSpPr;
     end;

     TCT_PieSer = class(TXPGBase)
protected
     FEG_SerShared: TEG_SerShared;
     FExplosion: TCT_UnsignedInt;
     FDPtXpgList: TCT_DPtXpgList;
     FDLbls: TCT_DLbls;
     FCat: TCT_AxDataSource;
     FVal: TCT_NumDataSource;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_PieSer);
     procedure CopyTo(AItem: TCT_PieSer);
     function  Create_Explosion: TCT_UnsignedInt;
     function  Create_DPtXpgList: TCT_DPtXpgList;
     function  Create_DLbls: TCT_DLbls;
     function  Create_Cat: TCT_AxDataSource;
     function  Create_Val: TCT_NumDataSource;
     function  Create_ExtLst: TCT_ExtensionList;

     property EG_SerShared: TEG_SerShared read FEG_SerShared;
     property Explosion: TCT_UnsignedInt read FExplosion;
     property DPtXpgList: TCT_DPtXpgList read FDPtXpgList;
     property DLbls: TCT_DLbls read FDLbls;
     property Cat: TCT_AxDataSource read FCat;
     property Val: TCT_NumDataSource read FVal;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_PieSerXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_PieSer;
public
     function  Add: TCT_PieSer;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_PieSerXpgList);
     procedure CopyTo(AItem: TCT_PieSerXpgList);
     property Items[Index: integer]: TCT_PieSer read GetItems; default;
     end;

     // Is required in order to keep excel happy.
     TCT_BarDir = class(TXPGBase)
protected
     FVal: TST_BarDir;
public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_BarDir);
     procedure CopyTo(AItem: TCT_BarDir);

     property Val: TST_BarDir read FVal write FVal;
     end;

     TCT_BarGrouping = class(TXPGBase)
protected
     FVal: TST_BarGrouping;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_BarGrouping);
     procedure CopyTo(AItem: TCT_BarGrouping);

     property Val: TST_BarGrouping read FVal write FVal;
     end;

     TCT_BarSer = class(TXPGBase)
protected
     FEG_SerShared: TEG_SerShared;
     FInvertIfNegative: TCT_Boolean;
     FPictureOptions: TCT_PictureOptions;
     FDPtXpgList: TCT_DPtXpgList;
     FDLbls: TCT_DLbls;
     FTrendlineXpgList: TCT_TrendlineXpgList;
     FErrBars: TCT_ErrBars;
     FCat: TCT_AxDataSource;
     FVal: TCT_NumDataSource;
     FShape: TCT_Shape;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_BarSer);
     procedure CopyTo(AItem: TCT_BarSer);
     function  Create_InvertIfNegative: TCT_Boolean;
     function  Create_PictureOptions: TCT_PictureOptions;
     function  Create_DPtXpgList: TCT_DPtXpgList;
     function  Create_DLbls: TCT_DLbls;
     function  Create_TrendlineXpgList: TCT_TrendlineXpgList;
     function  Create_ErrBars: TCT_ErrBars;
     function  Create_Cat: TCT_AxDataSource;
     function  Create_Val: TCT_NumDataSource;
     function  Create_Shape: TCT_Shape;
     function  Create_ExtLst: TCT_ExtensionList;

     property EG_SerShared: TEG_SerShared read FEG_SerShared;
     property InvertIfNegative: TCT_Boolean read FInvertIfNegative;
     property PictureOptions: TCT_PictureOptions read FPictureOptions;
     property DPtXpgList: TCT_DPtXpgList read FDPtXpgList;
     property DLbls: TCT_DLbls read FDLbls;
     property TrendlineXpgList: TCT_TrendlineXpgList read FTrendlineXpgList;
     property ErrBars: TCT_ErrBars read FErrBars;
     property Cat: TCT_AxDataSource read FCat;
     property Val: TCT_NumDataSource read FVal;
     property Shape: TCT_Shape read FShape;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_BarSerXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_BarSer;
public
     function  Add: TCT_BarSer;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_BarSerXpgList);
     procedure CopyTo(AItem: TCT_BarSerXpgList);
     property Items[Index: integer]: TCT_BarSer read GetItems; default;
     end;

     TCT_SurfaceSer = class(TXPGBase)
protected
     FEG_SerShared: TEG_SerShared;
     FCat: TCT_AxDataSource;
     FVal: TCT_NumDataSource;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_SurfaceSer);
     procedure CopyTo(AItem: TCT_SurfaceSer);
     function  Create_Cat: TCT_AxDataSource;
     function  Create_Val: TCT_NumDataSource;
     function  Create_ExtLst: TCT_ExtensionList;

     property EG_SerShared: TEG_SerShared read FEG_SerShared;
     property Cat: TCT_AxDataSource read FCat;
     property Val: TCT_NumDataSource read FVal;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_SurfaceSerXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_SurfaceSer;
public
     function  Add: TCT_SurfaceSer;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_SurfaceSerXpgList);
     procedure CopyTo(AItem: TCT_SurfaceSerXpgList);
     property Items[Index: integer]: TCT_SurfaceSer read GetItems; default;
     end;

     TCT_BandFmts = class(TXPGBase)
protected
     FBandFmtXpgList: TCT_BandFmtXpgList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_BandFmts);
     procedure CopyTo(AItem: TCT_BandFmts);
     function  Create_BandFmtXpgList: TCT_BandFmtXpgList;

     property BandFmtXpgList: TCT_BandFmtXpgList read FBandFmtXpgList;
     end;

     TCT_Scaling = class(TXPGBase)
protected
     FLogBase: TCT_LogBase;
     FOrientation: TCT_Orientation;
     FMax: TCT_Double;
     FMin: TCT_Double;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Scaling);
     procedure CopyTo(AItem: TCT_Scaling);
     function  Create_LogBase: TCT_LogBase;
     function  Create_Orientation: TCT_Orientation;
     function  Create_Max: TCT_Double;
     function  Create_Min: TCT_Double;
     function  Create_ExtLst: TCT_ExtensionList;

     property LogBase: TCT_LogBase read FLogBase;
     property Orientation: TCT_Orientation read FOrientation;
     property Max: TCT_Double read FMax;
     property Min: TCT_Double read FMin;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_AxPos = class(TXPGBase)
protected
     FVal: TST_AxPos;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_AxPos);
     procedure CopyTo(AItem: TCT_AxPos);

     property Val: TST_AxPos read FVal write FVal;
     end;

     TCT_Title = class(TXPGBase)
protected
     FTx: TCT_Tx;
     FLayout: TCT_Layout;
     FOverlay: TCT_Boolean;
     FSpPr: TCT_ShapeProperties;
     FTxPr: TCT_TextBody;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Title);
     procedure CopyTo(AItem: TCT_Title);
     function  Create_Tx: TCT_Tx;
     function  Create_Layout: TCT_Layout;
     function  Create_Overlay: TCT_Boolean;
     function  Create_SpPr: TCT_ShapeProperties;
     function  Create_TxPr: TCT_TextBody;
     function  Create_ExtLst: TCT_ExtensionList;

     property Tx: TCT_Tx read FTx;
     property Layout: TCT_Layout read FLayout;
     property Overlay: TCT_Boolean read FOverlay;
     property SpPr: TCT_ShapeProperties read FSpPr;
     property TxPr: TCT_TextBody read FTxPr;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_TickMark = class(TXPGBase)
protected
     FVal: TST_TickMark;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_TickMark);
     procedure CopyTo(AItem: TCT_TickMark);

     property Val: TST_TickMark read FVal write FVal;
     end;

     TCT_TickLblPos = class(TXPGBase)
protected
     FVal: TST_TickLblPos;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_TickLblPos);
     procedure CopyTo(AItem: TCT_TickLblPos);

     property Val: TST_TickLblPos read FVal write FVal;
     end;

     TCT_Crosses = class(TXPGBase)
protected
     FVal: TST_Crosses;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Crosses);
     procedure CopyTo(AItem: TCT_Crosses);

     property Val: TST_Crosses read FVal write FVal;
     end;

     TCT_BuiltInUnit = class(TXPGBase)
protected
     FVal: TST_BuiltInUnit;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_BuiltInUnit);
     procedure CopyTo(AItem: TCT_BuiltInUnit);

     property Val: TST_BuiltInUnit read FVal write FVal;
     end;

     TCT_DispUnitsLbl = class(TXPGBase)
protected
     FLayout: TCT_Layout;
     FTx: TCT_Tx;
     FSpPr: TCT_ShapeProperties;
     FTxPr: TCT_TextBody;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_DispUnitsLbl);
     procedure CopyTo(AItem: TCT_DispUnitsLbl);
     function  Create_Layout: TCT_Layout;
     function  Create_Tx: TCT_Tx;
     function  Create_SpPr: TCT_ShapeProperties;
     function  Create_TxPr: TCT_TextBody;

     property Layout: TCT_Layout read FLayout;
     property Tx: TCT_Tx read FTx;
     property SpPr: TCT_ShapeProperties read FSpPr;
     property TxPr: TCT_TextBody read FTxPr;
     end;

     TEG_AreaChartShared = class(TXPGBase)
protected
     FGrouping: TCT_Grouping;
     FVaryColors: TCT_Boolean;
     FSer: TCT_AreaSer;
     FDLbls: TCT_DLbls;
     FDropLines: TCT_ChartLines;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TEG_AreaChartShared);
     procedure CopyTo(AItem: TEG_AreaChartShared);
     function  Create_Grouping: TCT_Grouping;
     function  Create_VaryColors: TCT_Boolean;
     function  Create_Ser: TCT_AreaSer;
     function  Create_DLbls: TCT_DLbls;
     function  Create_DropLines: TCT_ChartLines;

     property Grouping: TCT_Grouping read FGrouping;
     property VaryColors: TCT_Boolean read FVaryColors;
     property Ser: TCT_AreaSer read FSer;
     property DLbls: TCT_DLbls read FDLbls;
     property DropLines: TCT_ChartLines read FDropLines;
     end;

     TEG_LineChartShared = class(TXPGBase)
protected
     FGrouping: TCT_Grouping;
     FVaryColors: TCT_Boolean;
     FSer: TCT_LineSerXpgList;
     FDLbls: TCT_DLbls;
     FDropLines: TCT_ChartLines;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TEG_LineChartShared);
     procedure CopyTo(AItem: TEG_LineChartShared);
     function  Create_Grouping: TCT_Grouping;
     function  Create_VaryColors: TCT_Boolean;
     function  Create_Ser: TCT_LineSerXpgList;
     function  Create_DLbls: TCT_DLbls;
     function  Create_DropLines: TCT_ChartLines;

     property Grouping: TCT_Grouping read FGrouping;
     property VaryColors: TCT_Boolean read FVaryColors;
     property Ser: TCT_LineSerXpgList read FSer;
     property DLbls: TCT_DLbls read FDLbls;
     property DropLines: TCT_ChartLines read FDropLines;
     end;

     TCT_UpDownBars = class(TXPGBase)
protected
     FGapWidth: TCT_GapAmount;
     FUpBars: TCT_UpDownBar;
     FDownBars: TCT_UpDownBar;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_UpDownBars);
     procedure CopyTo(AItem: TCT_UpDownBars);
     function  Create_GapWidth: TCT_GapAmount;
     function  Create_UpBars: TCT_UpDownBar;
     function  Create_DownBars: TCT_UpDownBar;
     function  Create_ExtLst: TCT_ExtensionList;

     property GapWidth: TCT_GapAmount read FGapWidth;
     property UpBars: TCT_UpDownBar read FUpBars;
     property DownBars: TCT_UpDownBar read FDownBars;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_RadarStyle = class(TXPGBase)
protected
     FVal: TST_RadarStyle;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_RadarStyle);
     procedure CopyTo(AItem: TCT_RadarStyle);

     property Val: TST_RadarStyle read FVal write FVal;
     end;

     TCT_RadarSer = class(TXPGBase)
protected
     FEG_SerShared: TEG_SerShared;
     FMarker: TCT_Marker;
     FDPtXpgList: TCT_DPtXpgList;
     FDLbls: TCT_DLbls;
     FCat: TCT_AxDataSource;
     FVal: TCT_NumDataSource;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_RadarSer);
     procedure CopyTo(AItem: TCT_RadarSer);
     function  Create_Marker: TCT_Marker;
     function  Create_DPtXpgList: TCT_DPtXpgList;
     function  Create_DLbls: TCT_DLbls;
     function  Create_Cat: TCT_AxDataSource;
     function  Create_Val: TCT_NumDataSource;
     function  Create_ExtLst: TCT_ExtensionList;

     property EG_SerShared: TEG_SerShared read FEG_SerShared;
     property Marker: TCT_Marker read FMarker;
     property DPtXpgList: TCT_DPtXpgList read FDPtXpgList;
     property DLbls: TCT_DLbls read FDLbls;
     property Cat: TCT_AxDataSource read FCat;
     property Val: TCT_NumDataSource read FVal;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_RadarSerXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_RadarSer;
public
     function  Add: TCT_RadarSer;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_RadarSerXpgList);
     procedure CopyTo(AItem: TCT_RadarSerXpgList);
     property Items[Index: integer]: TCT_RadarSer read GetItems; default;
     end;

     TCT_ScatterStyle = class(TXPGBase)
protected
     FVal: TST_ScatterStyle;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_ScatterStyle);
     procedure CopyTo(AItem: TCT_ScatterStyle);

     property Val: TST_ScatterStyle read FVal write FVal;
     end;

     TCT_ScatterSer = class(TXPGBase)
protected
     FEG_SerShared: TEG_SerShared;
     FMarker: TCT_Marker;
     FDPtXpgList: TCT_DPtXpgList;
     FDLbls: TCT_DLbls;
     FTrendlineXpgList: TCT_TrendlineXpgList;
     FErrBarsXpgList: TCT_ErrBarsXpgList;
     FXVal: TCT_AxDataSource;
     FYVal: TCT_NumDataSource;
     FSmooth: TCT_Boolean;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_ScatterSer);
     procedure CopyTo(AItem: TCT_ScatterSer);
     function  Create_Marker: TCT_Marker;
     function  Create_DPtXpgList: TCT_DPtXpgList;
     function  Create_DLbls: TCT_DLbls;
     function  Create_TrendlineXpgList: TCT_TrendlineXpgList;
     function  Create_ErrBarsXpgList: TCT_ErrBarsXpgList;
     function  Create_XVal: TCT_AxDataSource;
     function  Create_YVal: TCT_NumDataSource;
     function  Create_Smooth: TCT_Boolean;
     function  Create_ExtLst: TCT_ExtensionList;

     property EG_SerShared: TEG_SerShared read FEG_SerShared;
     property Marker: TCT_Marker read FMarker;
     property DPtXpgList: TCT_DPtXpgList read FDPtXpgList;
     property DLbls: TCT_DLbls read FDLbls;
     property TrendlineXpgList: TCT_TrendlineXpgList read FTrendlineXpgList;
     property ErrBarsXpgList: TCT_ErrBarsXpgList read FErrBarsXpgList;
     property XVal: TCT_AxDataSource read FXVal;
     property YVal: TCT_NumDataSource read FYVal;
     property Smooth: TCT_Boolean read FSmooth;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_ScatterSerXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_ScatterSer;
public
     function  Add: TCT_ScatterSer;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_ScatterSerXpgList);
     procedure CopyTo(AItem: TCT_ScatterSerXpgList);
     property Items[Index: integer]: TCT_ScatterSer read GetItems; default;
     end;

     TEG_PieChartShared = class(TXPGBase)
protected
     FVaryColors: TCT_Boolean;
     FSer: TCT_PieSer;
     FDLbls: TCT_DLbls;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TEG_PieChartShared);
     procedure CopyTo(AItem: TEG_PieChartShared);
     function  Create_VaryColors: TCT_Boolean;
     function  Create_Ser: TCT_PieSer;
     function  Create_DLbls: TCT_DLbls;

     property VaryColors: TCT_Boolean read FVaryColors;
     property Ser: TCT_PieSer read FSer;
     property DLbls: TCT_DLbls read FDLbls;
     end;

     TCT_FirstSliceAng = class(TXPGBase)
protected
     FVal: integer;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_FirstSliceAng);
     procedure CopyTo(AItem: TCT_FirstSliceAng);

     property Val: integer read FVal write FVal;
     end;

     TCT_HoleSize = class(TXPGBase)
protected
     FVal: integer;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_HoleSize);
     procedure CopyTo(AItem: TCT_HoleSize);

     property Val: integer read FVal write FVal;
     end;

     TEG_BarChartShared = class(TXPGBase)
protected
     FBarDir: TCT_BarDir;
     FGrouping: TCT_BarGrouping;
     FVaryColors: TCT_Boolean;
     FSeries: TXPGBaseList;
     FDLbls: TCT_DLbls;
public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TEG_BarChartShared);
     procedure CopyTo(AItem: TEG_BarChartShared);
     function  Create_BarDir: TCT_BarDir;
     function  Create_Grouping: TCT_BarGrouping;
     function  Create_VaryColors: TCT_Boolean;
     function  Create_DLbls: TCT_DLbls;

     property BarDir: TCT_BarDir read FBarDir;
     property Grouping: TCT_BarGrouping read FGrouping;
     property VaryColors: TCT_Boolean read FVaryColors;
     property Series: TXPGBaseList read FSeries;
     property DLbls: TCT_DLbls read FDLbls;
     end;

     TCT_Overlap = class(TXPGBase)
protected
     FVal: integer;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Overlap);
     procedure CopyTo(AItem: TCT_Overlap);

     property Val: integer read FVal write FVal;
     end;

     TCT_OfPieType = class(TXPGBase)
protected
     FVal: TST_OfPieType;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_OfPieType);
     procedure CopyTo(AItem: TCT_OfPieType);

     property Val: TST_OfPieType read FVal write FVal;
     end;

     TCT_SplitType = class(TXPGBase)
protected
     FVal: TST_SplitType;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_SplitType);
     procedure CopyTo(AItem: TCT_SplitType);

     property Val: TST_SplitType read FVal write FVal;
     end;

     TCT_CustSplit = class(TXPGBase)
protected
     FSecondPiePtXpgList: TCT_UnsignedIntXpgList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_CustSplit);
     procedure CopyTo(AItem: TCT_CustSplit);
     function  Create_SecondPiePtXpgList: TCT_UnsignedIntXpgList;

     property SecondPiePtXpgList: TCT_UnsignedIntXpgList read FSecondPiePtXpgList;
     end;

     TCT_SecondPieSize = class(TXPGBase)
protected
     FVal: integer;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_SecondPieSize);
     procedure CopyTo(AItem: TCT_SecondPieSize);

     property Val: integer read FVal write FVal;
     end;

     TEG_SurfaceChartShared = class(TXPGBase)
protected
     FWireframe: TCT_Boolean;
     FSer: TCT_SurfaceSer;
     FBandFmts: TCT_BandFmts;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TEG_SurfaceChartShared);
     procedure CopyTo(AItem: TEG_SurfaceChartShared);
     function  Create_Wireframe: TCT_Boolean;
     function  Create_Ser: TCT_SurfaceSer;
     function  Create_BandFmts: TCT_BandFmts;

     property Wireframe: TCT_Boolean read FWireframe;
     property Ser: TCT_SurfaceSer read FSer;
     property BandFmts: TCT_BandFmts read FBandFmts;
     end;

     TCT_BubbleSer = class(TXPGBase)
protected
     FEG_SerShared: TEG_SerShared;
     FInvertIfNegative: TCT_Boolean;
     FDPtXpgList: TCT_DPtXpgList;
     FDLbls: TCT_DLbls;
     FTrendlineXpgList: TCT_TrendlineXpgList;
     FErrBarsXpgList: TCT_ErrBarsXpgList;
     FXVal: TCT_AxDataSource;
     FYVal: TCT_NumDataSource;
     FBubbleSize: TCT_NumDataSource;
     FBubble3D: TCT_Boolean;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_BubbleSer);
     procedure CopyTo(AItem: TCT_BubbleSer);
     function  Create_InvertIfNegative: TCT_Boolean;
     function  Create_DPtXpgList: TCT_DPtXpgList;
     function  Create_DLbls: TCT_DLbls;
     function  Create_TrendlineXpgList: TCT_TrendlineXpgList;
     function  Create_ErrBarsXpgList: TCT_ErrBarsXpgList;
     function  Create_XVal: TCT_AxDataSource;
     function  Create_YVal: TCT_NumDataSource;
     function  Create_BubbleSize: TCT_NumDataSource;
     function  Create_Bubble3D: TCT_Boolean;
     function  Create_ExtLst: TCT_ExtensionList;

     property EG_SerShared: TEG_SerShared read FEG_SerShared;
     property InvertIfNegative: TCT_Boolean read FInvertIfNegative;
     property DPtXpgList: TCT_DPtXpgList read FDPtXpgList;
     property DLbls: TCT_DLbls read FDLbls;
     property TrendlineXpgList: TCT_TrendlineXpgList read FTrendlineXpgList;
     property ErrBarsXpgList: TCT_ErrBarsXpgList read FErrBarsXpgList;
     property XVal: TCT_AxDataSource read FXVal;
     property YVal: TCT_NumDataSource read FYVal;
     property BubbleSize: TCT_NumDataSource read FBubbleSize;
     property Bubble3D: TCT_Boolean read FBubble3D;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_BubbleSerXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_BubbleSer;
public
     function  Add: TCT_BubbleSer;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_BubbleSerXpgList);
     procedure CopyTo(AItem: TCT_BubbleSerXpgList);
     property Items[Index: integer]: TCT_BubbleSer read GetItems; default;
     end;

     TCT_BubbleScale = class(TXPGBase)
protected
     FVal: integer;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_BubbleScale);
     procedure CopyTo(AItem: TCT_BubbleScale);

     property Val: integer read FVal write FVal;
     end;

     TCT_SizeRepresents = class(TXPGBase)
protected
     FVal: TST_SizeRepresents;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_SizeRepresents);
     procedure CopyTo(AItem: TCT_SizeRepresents);

     property Val: TST_SizeRepresents read FVal write FVal;
     end;

     TEG_AxShared = class(TXPGBase)
protected
     FAxId: TCT_UnsignedInt;
     FScaling: TCT_Scaling;
     FDelete: TCT_Boolean;
     FAxPos: TCT_AxPos;
     FMajorGridlines: TCT_ChartLines;
     FMinorGridlines: TCT_ChartLines;
     FTitle: TCT_Title;
     FNumFmt: TCT_NumFmt;
     FMajorTickMark: TCT_TickMark;
     FMinorTickMark: TCT_TickMark;
     FTickLblPos: TCT_TickLblPos;
     FSpPr: TCT_ShapeProperties;
     FTxPr: TCT_TextBody;
     FCrossAx: TCT_UnsignedInt;
     FCrosses: TCT_Crosses;
     FCrossesAt: TCT_Double;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TEG_AxShared);
     procedure CopyTo(AItem: TEG_AxShared);
     function  Create_Delete: TCT_Boolean;
     function  Create_MajorGridlines: TCT_ChartLines;
     function  Create_MinorGridlines: TCT_ChartLines;
     function  Create_Title: TCT_Title;
     function  Create_NumFmt: TCT_NumFmt;
     function  Create_MajorTickMark: TCT_TickMark;
     function  Create_MinorTickMark: TCT_TickMark;
     function  Create_TickLblPos: TCT_TickLblPos;
     function  Create_SpPr: TCT_ShapeProperties;
     function  Create_TxPr: TCT_TextBody;
     function  Create_Crosses: TCT_Crosses;
     function  Create_CrossesAt: TCT_Double;

     property AxId: TCT_UnsignedInt read FAxId;
     property Scaling: TCT_Scaling read FScaling;
     property Delete: TCT_Boolean read FDelete;
     property AxPos: TCT_AxPos read FAxPos;
     property MajorGridlines: TCT_ChartLines read FMajorGridlines;
     property MinorGridlines: TCT_ChartLines read FMinorGridlines;
     property Title: TCT_Title read FTitle;
     property NumFmt: TCT_NumFmt read FNumFmt;
     property MajorTickMark: TCT_TickMark read FMajorTickMark;
     property MinorTickMark: TCT_TickMark read FMinorTickMark;
     property TickLblPos: TCT_TickLblPos read FTickLblPos;
     property SpPr: TCT_ShapeProperties read FSpPr;
     property TxPr: TCT_TextBody read FTxPr;
     property CrossAx: TCT_UnsignedInt read FCrossAx;
     property Crosses: TCT_Crosses read FCrosses;
     property CrossesAt: TCT_Double read FCrossesAt;
     end;

     TCT_CrossBetween = class(TXPGBase)
protected
     FVal: TST_CrossBetween;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_CrossBetween);
     procedure CopyTo(AItem: TCT_CrossBetween);

     property Val: TST_CrossBetween read FVal write FVal;
     end;

     TCT_AxisUnit = class(TXPGBase)
protected
     FVal: double;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_AxisUnit);
     procedure CopyTo(AItem: TCT_AxisUnit);

     property Val: double read FVal write FVal;
     end;

     TCT_DispUnits = class(TXPGBase)
protected
     FCustUnit: TCT_Double;
     FBuiltInUnit: TCT_BuiltInUnit;
     FDispUnitsLbl: TCT_DispUnitsLbl;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_DispUnits);
     procedure CopyTo(AItem: TCT_DispUnits);
     function  Create_CustUnit: TCT_Double;
     function  Create_BuiltInUnit: TCT_BuiltInUnit;
     function  Create_DispUnitsLbl: TCT_DispUnitsLbl;
     function  Create_ExtLst: TCT_ExtensionList;

     property CustUnit: TCT_Double read FCustUnit;
     property BuiltInUnit: TCT_BuiltInUnit read FBuiltInUnit;
     property DispUnitsLbl: TCT_DispUnitsLbl read FDispUnitsLbl;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_LblAlgn = class(TXPGBase)
protected
     FVal: TST_LblAlgn;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_LblAlgn);
     procedure CopyTo(AItem: TCT_LblAlgn);

     property Val: TST_LblAlgn read FVal write FVal;
     end;

     TCT_LblOffset = class(TXPGBase)
protected
     FVal: integer;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_LblOffset);
     procedure CopyTo(AItem: TCT_LblOffset);

     property Val: integer read FVal write FVal;
     end;

     TCT_Skip = class(TXPGBase)
protected
     FVal: integer;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Skip);
     procedure CopyTo(AItem: TCT_Skip);

     property Val: integer read FVal write FVal;
     end;

     TCT_TimeUnit = class(TXPGBase)
protected
     FVal: TST_TimeUnit;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_TimeUnit);
     procedure CopyTo(AItem: TCT_TimeUnit);

     property Val: TST_TimeUnit read FVal write FVal;
     end;

     TEG_LegendEntryData = class(TXPGBase)
protected
     FTxPr: TCT_TextBody;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TEG_LegendEntryData);
     procedure CopyTo(AItem: TEG_LegendEntryData);
     function  Create_TxPr: TCT_TextBody;

     property TxPr: TCT_TextBody read FTxPr;
     end;

     TCT_ColorMapping = class(TXPGBase)
protected
     FBg1: TST_ColorSchemeIndex;
     FTx1: TST_ColorSchemeIndex;
     FBg2: TST_ColorSchemeIndex;
     FTx2: TST_ColorSchemeIndex;
     FAccent1: TST_ColorSchemeIndex;
     FAccent2: TST_ColorSchemeIndex;
     FAccent3: TST_ColorSchemeIndex;
     FAccent4: TST_ColorSchemeIndex;
     FAccent5: TST_ColorSchemeIndex;
     FAccent6: TST_ColorSchemeIndex;
     FHlink: TST_ColorSchemeIndex;
     FFolHlink: TST_ColorSchemeIndex;
     FA_ExtLst: TCT_OfficeArtExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_ColorMapping);
     procedure CopyTo(AItem: TCT_ColorMapping);
     function  Create_A_ExtLst: TCT_OfficeArtExtensionList;

     property Bg1: TST_ColorSchemeIndex read FBg1 write FBg1;
     property Tx1: TST_ColorSchemeIndex read FTx1 write FTx1;
     property Bg2: TST_ColorSchemeIndex read FBg2 write FBg2;
     property Tx2: TST_ColorSchemeIndex read FTx2 write FTx2;
     property Accent1: TST_ColorSchemeIndex read FAccent1 write FAccent1;
     property Accent2: TST_ColorSchemeIndex read FAccent2 write FAccent2;
     property Accent3: TST_ColorSchemeIndex read FAccent3 write FAccent3;
     property Accent4: TST_ColorSchemeIndex read FAccent4 write FAccent4;
     property Accent5: TST_ColorSchemeIndex read FAccent5 write FAccent5;
     property Accent6: TST_ColorSchemeIndex read FAccent6 write FAccent6;
     property Hlink: TST_ColorSchemeIndex read FHlink write FHlink;
     property FolHlink: TST_ColorSchemeIndex read FFolHlink write FFolHlink;
     property A_ExtLst: TCT_OfficeArtExtensionList read FA_ExtLst;
     end;

     TCT_PivotFmt = class(TXPGBase)
protected
     FIdx: TCT_UnsignedInt;
     FSpPr: TCT_ShapeProperties;
     FTxPr: TCT_TextBody;
     FMarker: TCT_Marker;
     FDLbl: TCT_DLbl;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_PivotFmt);
     procedure CopyTo(AItem: TCT_PivotFmt);
     function  Create_Idx: TCT_UnsignedInt;
     function  Create_SpPr: TCT_ShapeProperties;
     function  Create_TxPr: TCT_TextBody;
     function  Create_Marker: TCT_Marker;
     function  Create_DLbl: TCT_DLbl;
     function  Create_ExtLst: TCT_ExtensionList;

     property Idx: TCT_UnsignedInt read FIdx;
     property SpPr: TCT_ShapeProperties read FSpPr;
     property TxPr: TCT_TextBody read FTxPr;
     property Marker: TCT_Marker read FMarker;
     property DLbl: TCT_DLbl read FDLbl;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_PivotFmtXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_PivotFmt;
public
     function  Add: TCT_PivotFmt;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_PivotFmtXpgList);
     procedure CopyTo(AItem: TCT_PivotFmtXpgList);
     property Items[Index: integer]: TCT_PivotFmt read GetItems; default;
     end;

     TCT_RotX = class(TXPGBase)
protected
     FVal: integer;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_RotX);
     procedure CopyTo(AItem: TCT_RotX);

     property Val: integer read FVal write FVal;
     end;

     TCT_HPercent = class(TXPGBase)
protected
     FVal: integer;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_HPercent);
     procedure CopyTo(AItem: TCT_HPercent);

     property Val: integer read FVal write FVal;
     end;

     TCT_RotY = class(TXPGBase)
protected
     FVal: integer;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_RotY);
     procedure CopyTo(AItem: TCT_RotY);

     property Val: integer read FVal write FVal;
     end;

     TCT_DepthPercent = class(TXPGBase)
protected
     FVal: integer;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_DepthPercent);
     procedure CopyTo(AItem: TCT_DepthPercent);

     property Val: integer read FVal write FVal;
     end;

     TCT_Perspective = class(TXPGBase)
protected
     FVal: integer;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Perspective);
     procedure CopyTo(AItem: TCT_Perspective);

     property Val: integer read FVal write FVal;
     end;

     TCT_AreaChart = class(TXPGBase)
protected
     FEG_AreaChartShared: TEG_AreaChartShared;
     FAxIdXpgList: TCT_UnsignedIntXpgList;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_AreaChart);
     procedure CopyTo(AItem: TCT_AreaChart);
     function  Create_AxIdXpgList: TCT_UnsignedIntXpgList;
     function  Create_ExtLst: TCT_ExtensionList;

     property EG_AreaChartShared: TEG_AreaChartShared read FEG_AreaChartShared;
     property AxIdXpgList: TCT_UnsignedIntXpgList read FAxIdXpgList;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_AreaChartXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_AreaChart;
public
     function  Add: TCT_AreaChart;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_AreaChartXpgList);
     procedure CopyTo(AItem: TCT_AreaChartXpgList);
     property Items[Index: integer]: TCT_AreaChart read GetItems; default;
     end;

     TCT_Area3DChart = class(TXPGBase)
protected
     FEG_AreaChartShared: TEG_AreaChartShared;
     FGapDepth: TCT_GapAmount;
     FAxIdXpgList: TCT_UnsignedIntXpgList;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Area3DChart);
     procedure CopyTo(AItem: TCT_Area3DChart);
     function  Create_GapDepth: TCT_GapAmount;
     function  Create_AxIdXpgList: TCT_UnsignedIntXpgList;
     function  Create_ExtLst: TCT_ExtensionList;

     property EG_AreaChartShared: TEG_AreaChartShared read FEG_AreaChartShared;
     property GapDepth: TCT_GapAmount read FGapDepth;
     property AxIdXpgList: TCT_UnsignedIntXpgList read FAxIdXpgList;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_Area3DChartXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_Area3DChart;
public
     function  Add: TCT_Area3DChart;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_Area3DChartXpgList);
     procedure CopyTo(AItem: TCT_Area3DChartXpgList);
     property Items[Index: integer]: TCT_Area3DChart read GetItems; default;
     end;

     TCT_LineChart = class(TXPGBase)
protected
     FEG_LineChartShared: TEG_LineChartShared;
     FHiLowLines: TCT_ChartLines;
     FUpDownBars: TCT_UpDownBars;
     FMarker: TCT_Boolean;
     FSmooth: TCT_Boolean;
     FAxIdXpgList: TCT_UnsignedIntXpgList;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_LineChart);
     procedure CopyTo(AItem: TCT_LineChart);
     function  Create_HiLowLines: TCT_ChartLines;
     function  Create_UpDownBars: TCT_UpDownBars;
     function  Create_Marker: TCT_Boolean;
     function  Create_Smooth: TCT_Boolean;
     function  Create_AxIdXpgList: TCT_UnsignedIntXpgList;
     function  Create_ExtLst: TCT_ExtensionList;

     property EG_LineChartShared: TEG_LineChartShared read FEG_LineChartShared;
     property HiLowLines: TCT_ChartLines read FHiLowLines;
     property UpDownBars: TCT_UpDownBars read FUpDownBars;
     property Marker: TCT_Boolean read FMarker;
     property Smooth: TCT_Boolean read FSmooth;
     property AxIdXpgList: TCT_UnsignedIntXpgList read FAxIdXpgList;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_LineChartXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_LineChart;
public
     function  Add: TCT_LineChart;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_LineChartXpgList);
     procedure CopyTo(AItem: TCT_LineChartXpgList);
     property Items[Index: integer]: TCT_LineChart read GetItems; default;
     end;

     TCT_Line3DChart = class(TXPGBase)
protected
     FEG_LineChartShared: TEG_LineChartShared;
     FGapDepth: TCT_GapAmount;
     FAxIdXpgList: TCT_UnsignedIntXpgList;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Line3DChart);
     procedure CopyTo(AItem: TCT_Line3DChart);
     function  Create_GapDepth: TCT_GapAmount;
     function  Create_AxIdXpgList: TCT_UnsignedIntXpgList;
     function  Create_ExtLst: TCT_ExtensionList;

     property EG_LineChartShared: TEG_LineChartShared read FEG_LineChartShared;
     property GapDepth: TCT_GapAmount read FGapDepth;
     property AxIdXpgList: TCT_UnsignedIntXpgList read FAxIdXpgList;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_Line3DChartXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_Line3DChart;
public
     function  Add: TCT_Line3DChart;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_Line3DChartXpgList);
     procedure CopyTo(AItem: TCT_Line3DChartXpgList);
     property Items[Index: integer]: TCT_Line3DChart read GetItems; default;
     end;

     TCT_StockChart = class(TXPGBase)
protected
     FSerXpgList: TCT_LineSerXpgList;
     FDLbls: TCT_DLbls;
     FDropLines: TCT_ChartLines;
     FHiLowLines: TCT_ChartLines;
     FUpDownBars: TCT_UpDownBars;
     FAxIdXpgList: TCT_UnsignedIntXpgList;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_StockChart);
     procedure CopyTo(AItem: TCT_StockChart);
     function  Create_SerXpgList: TCT_LineSerXpgList;
     function  Create_DLbls: TCT_DLbls;
     function  Create_DropLines: TCT_ChartLines;
     function  Create_HiLowLines: TCT_ChartLines;
     function  Create_UpDownBars: TCT_UpDownBars;
     function  Create_AxIdXpgList: TCT_UnsignedIntXpgList;
     function  Create_ExtLst: TCT_ExtensionList;

     property SerXpgList: TCT_LineSerXpgList read FSerXpgList;
     property DLbls: TCT_DLbls read FDLbls;
     property DropLines: TCT_ChartLines read FDropLines;
     property HiLowLines: TCT_ChartLines read FHiLowLines;
     property UpDownBars: TCT_UpDownBars read FUpDownBars;
     property AxIdXpgList: TCT_UnsignedIntXpgList read FAxIdXpgList;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_StockChartXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_StockChart;
public
     function  Add: TCT_StockChart;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_StockChartXpgList);
     procedure CopyTo(AItem: TCT_StockChartXpgList);
     property Items[Index: integer]: TCT_StockChart read GetItems; default;
     end;

     TCT_RadarChart = class(TXPGBase)
protected
     FRadarStyle: TCT_RadarStyle;
     FVaryColors: TCT_Boolean;
     FSerXpgList: TCT_RadarSerXpgList;
     FDLbls: TCT_DLbls;
     FAxIdXpgList: TCT_UnsignedIntXpgList;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_RadarChart);
     procedure CopyTo(AItem: TCT_RadarChart);
     function  Create_RadarStyle: TCT_RadarStyle;
     function  Create_VaryColors: TCT_Boolean;
     function  Create_SerXpgList: TCT_RadarSerXpgList;
     function  Create_DLbls: TCT_DLbls;
     function  Create_AxIdXpgList: TCT_UnsignedIntXpgList;
     function  Create_ExtLst: TCT_ExtensionList;

     property RadarStyle: TCT_RadarStyle read FRadarStyle;
     property VaryColors: TCT_Boolean read FVaryColors;
     property SerXpgList: TCT_RadarSerXpgList read FSerXpgList;
     property DLbls: TCT_DLbls read FDLbls;
     property AxIdXpgList: TCT_UnsignedIntXpgList read FAxIdXpgList;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_RadarChartXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_RadarChart;
public
     function  Add: TCT_RadarChart;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_RadarChartXpgList);
     procedure CopyTo(AItem: TCT_RadarChartXpgList);
     property Items[Index: integer]: TCT_RadarChart read GetItems; default;
     end;

     TCT_ScatterChart = class(TXPGBase)
protected
     FScatterStyle: TCT_ScatterStyle;
     FVaryColors: TCT_Boolean;
     FSerXpgList: TCT_ScatterSerXpgList;
     FDLbls: TCT_DLbls;
     FAxIdXpgList: TCT_UnsignedIntXpgList;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_ScatterChart);
     procedure CopyTo(AItem: TCT_ScatterChart);
     function  Create_ScatterStyle: TCT_ScatterStyle;
     function  Create_VaryColors: TCT_Boolean;
     function  Create_SerXpgList: TCT_ScatterSerXpgList;
     function  Create_DLbls: TCT_DLbls;
     function  Create_AxIdXpgList: TCT_UnsignedIntXpgList;
     function  Create_ExtLst: TCT_ExtensionList;

     property ScatterStyle: TCT_ScatterStyle read FScatterStyle;
     property VaryColors: TCT_Boolean read FVaryColors;
     property SerXpgList: TCT_ScatterSerXpgList read FSerXpgList;
     property DLbls: TCT_DLbls read FDLbls;
     property AxIdXpgList: TCT_UnsignedIntXpgList read FAxIdXpgList;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_ScatterChartXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_ScatterChart;
public
     function  Add: TCT_ScatterChart;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_ScatterChartXpgList);
     procedure CopyTo(AItem: TCT_ScatterChartXpgList);
     property Items[Index: integer]: TCT_ScatterChart read GetItems; default;
     end;

     TCT_PieChart = class(TXPGBase)
protected
     FEG_PieChartShared: TEG_PieChartShared;
     FFirstSliceAng: TCT_FirstSliceAng;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_PieChart);
     procedure CopyTo(AItem: TCT_PieChart);
     function  Create_FirstSliceAng: TCT_FirstSliceAng;
     function  Create_ExtLst: TCT_ExtensionList;

     property EG_PieChartShared: TEG_PieChartShared read FEG_PieChartShared;
     property FirstSliceAng: TCT_FirstSliceAng read FFirstSliceAng;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_PieChartXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_PieChart;
public
     function  Add: TCT_PieChart;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_PieChartXpgList);
     procedure CopyTo(AItem: TCT_PieChartXpgList);
     property Items[Index: integer]: TCT_PieChart read GetItems; default;
     end;

     TCT_Pie3DChart = class(TXPGBase)
protected
     FEG_PieChartShared: TEG_PieChartShared;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Pie3DChart);
     procedure CopyTo(AItem: TCT_Pie3DChart);
     function  Create_ExtLst: TCT_ExtensionList;

     property EG_PieChartShared: TEG_PieChartShared read FEG_PieChartShared;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_Pie3DChartXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_Pie3DChart;
public
     function  Add: TCT_Pie3DChart;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_Pie3DChartXpgList);
     procedure CopyTo(AItem: TCT_Pie3DChartXpgList);
     property Items[Index: integer]: TCT_Pie3DChart read GetItems; default;
     end;

     TCT_DoughnutChart = class(TXPGBase)
protected
     FEG_PieChartShared: TEG_PieChartShared;
     FFirstSliceAng: TCT_FirstSliceAng;
     FHoleSize: TCT_HoleSize;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_DoughnutChart);
     procedure CopyTo(AItem: TCT_DoughnutChart);
     function  Create_FirstSliceAng: TCT_FirstSliceAng;
     function  Create_HoleSize: TCT_HoleSize;
     function  Create_ExtLst: TCT_ExtensionList;

     property EG_PieChartShared: TEG_PieChartShared read FEG_PieChartShared;
     property FirstSliceAng: TCT_FirstSliceAng read FFirstSliceAng;
     property HoleSize: TCT_HoleSize read FHoleSize;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_DoughnutChartXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_DoughnutChart;
public
     function  Add: TCT_DoughnutChart;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_DoughnutChartXpgList);
     procedure CopyTo(AItem: TCT_DoughnutChartXpgList);
     property Items[Index: integer]: TCT_DoughnutChart read GetItems; default;
     end;

     TCT_BarChart = class(TXPGBase)
protected
     FEG_BarChartShared: TEG_BarChartShared;
     FGapWidth: TCT_GapAmount;
     FOverlap: TCT_Overlap;
     FSerLinesXpgList: TCT_ChartLinesXpgList;
     FAxIdXpgList: TCT_UnsignedIntXpgList;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_BarChart);
     procedure CopyTo(AItem: TCT_BarChart);
     function  Create_GapWidth: TCT_GapAmount;
     function  Create_Overlap: TCT_Overlap;
     function  Create_SerLinesXpgList: TCT_ChartLinesXpgList;
     function  Create_AxIdXpgList: TCT_UnsignedIntXpgList;
     function  Create_ExtLst: TCT_ExtensionList;

     property EG_BarChartShared: TEG_BarChartShared read FEG_BarChartShared;
     property GapWidth: TCT_GapAmount read FGapWidth;
     property Overlap: TCT_Overlap read FOverlap;
     property SerLinesXpgList: TCT_ChartLinesXpgList read FSerLinesXpgList;
     property AxIdXpgList: TCT_UnsignedIntXpgList read FAxIdXpgList;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_BarChartXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_BarChart;
public
     function  Add: TCT_BarChart;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_BarChartXpgList);
     procedure CopyTo(AItem: TCT_BarChartXpgList);
     property Items[Index: integer]: TCT_BarChart read GetItems; default;
     end;

     TCT_Bar3DChart = class(TXPGBase)
protected
     FEG_BarChartShared: TEG_BarChartShared;
     FGapWidth: TCT_GapAmount;
     FGapDepth: TCT_GapAmount;
     FShape: TCT_Shape_Chart;
     FAxIdXpgList: TCT_UnsignedIntXpgList;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Bar3DChart);
     procedure CopyTo(AItem: TCT_Bar3DChart);
     function  Create_GapWidth: TCT_GapAmount;
     function  Create_GapDepth: TCT_GapAmount;
     function  Create_Shape: TCT_Shape_Chart;
     function  Create_AxIdXpgList: TCT_UnsignedIntXpgList;
     function  Create_ExtLst: TCT_ExtensionList;

     property EG_BarChartShared: TEG_BarChartShared read FEG_BarChartShared;
     property GapWidth: TCT_GapAmount read FGapWidth;
     property GapDepth: TCT_GapAmount read FGapDepth;
     property Shape: TCT_Shape_Chart read FShape;
     property AxIdXpgList: TCT_UnsignedIntXpgList read FAxIdXpgList;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_Bar3DChartXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_Bar3DChart;
public
     function  Add: TCT_Bar3DChart;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_Bar3DChartXpgList);
     procedure CopyTo(AItem: TCT_Bar3DChartXpgList);
     property Items[Index: integer]: TCT_Bar3DChart read GetItems; default;
     end;

     TCT_OfPieChart = class(TXPGBase)
protected
     FOfPieType: TCT_OfPieType;
     FEG_PieChartShared: TEG_PieChartShared;
     FGapWidth: TCT_GapAmount;
     FSplitType: TCT_SplitType;
     FSplitPos: TCT_Double;
     FCustSplit: TCT_CustSplit;
     FSecondPieSize: TCT_SecondPieSize;
     FSerLinesXpgList: TCT_ChartLinesXpgList;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_OfPieChart);
     procedure CopyTo(AItem: TCT_OfPieChart);
     function  Create_OfPieType: TCT_OfPieType;
     function  Create_GapWidth: TCT_GapAmount;
     function  Create_SplitType: TCT_SplitType;
     function  Create_SplitPos: TCT_Double;
     function  Create_CustSplit: TCT_CustSplit;
     function  Create_SecondPieSize: TCT_SecondPieSize;
     function  Create_SerLinesXpgList: TCT_ChartLinesXpgList;
     function  Create_ExtLst: TCT_ExtensionList;

     property OfPieType: TCT_OfPieType read FOfPieType;
     property EG_PieChartShared: TEG_PieChartShared read FEG_PieChartShared;
     property GapWidth: TCT_GapAmount read FGapWidth;
     property SplitType: TCT_SplitType read FSplitType;
     property SplitPos: TCT_Double read FSplitPos;
     property CustSplit: TCT_CustSplit read FCustSplit;
     property SecondPieSize: TCT_SecondPieSize read FSecondPieSize;
     property SerLinesXpgList: TCT_ChartLinesXpgList read FSerLinesXpgList;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_OfPieChartXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_OfPieChart;
public
     function  Add: TCT_OfPieChart;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_OfPieChartXpgList);
     procedure CopyTo(AItem: TCT_OfPieChartXpgList);
     property Items[Index: integer]: TCT_OfPieChart read GetItems; default;
     end;

     TCT_SurfaceChart = class(TXPGBase)
protected
     FEG_SurfaceChartShared: TEG_SurfaceChartShared;
     FAxIdXpgList: TCT_UnsignedIntXpgList;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_SurfaceChart);
     procedure CopyTo(AItem: TCT_SurfaceChart);
     function  Create_AxIdXpgList: TCT_UnsignedIntXpgList;
     function  Create_ExtLst: TCT_ExtensionList;

     property EG_SurfaceChartShared: TEG_SurfaceChartShared read FEG_SurfaceChartShared;
     property AxIdXpgList: TCT_UnsignedIntXpgList read FAxIdXpgList;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_SurfaceChartXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_SurfaceChart;
public
     function  Add: TCT_SurfaceChart;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_SurfaceChartXpgList);
     procedure CopyTo(AItem: TCT_SurfaceChartXpgList);
     property Items[Index: integer]: TCT_SurfaceChart read GetItems; default;
     end;

     TCT_Surface3DChart = class(TXPGBase)
protected
     FEG_SurfaceChartShared: TEG_SurfaceChartShared;
     FAxIdXpgList: TCT_UnsignedIntXpgList;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Surface3DChart);
     procedure CopyTo(AItem: TCT_Surface3DChart);
     function  Create_AxIdXpgList: TCT_UnsignedIntXpgList;
     function  Create_ExtLst: TCT_ExtensionList;

     property EG_SurfaceChartShared: TEG_SurfaceChartShared read FEG_SurfaceChartShared;
     property AxIdXpgList: TCT_UnsignedIntXpgList read FAxIdXpgList;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_Surface3DChartXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_Surface3DChart;
public
     function  Add: TCT_Surface3DChart;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_Surface3DChartXpgList);
     procedure CopyTo(AItem: TCT_Surface3DChartXpgList);
     property Items[Index: integer]: TCT_Surface3DChart read GetItems; default;
     end;

     TCT_BubbleChart = class(TXPGBase)
protected
     FVaryColors: TCT_Boolean;
     FSerXpgList: TCT_BubbleSerXpgList;
     FDLbls: TCT_DLbls;
     FBubble3D: TCT_Boolean;
     FBubbleScale: TCT_BubbleScale;
     FShowNegBubbles: TCT_Boolean;
     FSizeRepresents: TCT_SizeRepresents;
     FAxIdXpgList: TCT_UnsignedIntXpgList;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_BubbleChart);
     procedure CopyTo(AItem: TCT_BubbleChart);
     function  Create_VaryColors: TCT_Boolean;
     function  Create_SerXpgList: TCT_BubbleSerXpgList;
     function  Create_DLbls: TCT_DLbls;
     function  Create_Bubble3D: TCT_Boolean;
     function  Create_BubbleScale: TCT_BubbleScale;
     function  Create_ShowNegBubbles: TCT_Boolean;
     function  Create_SizeRepresents: TCT_SizeRepresents;
     function  Create_AxIdXpgList: TCT_UnsignedIntXpgList;
     function  Create_ExtLst: TCT_ExtensionList;

     property VaryColors: TCT_Boolean read FVaryColors;
     property SerXpgList: TCT_BubbleSerXpgList read FSerXpgList;
     property DLbls: TCT_DLbls read FDLbls;
     property Bubble3D: TCT_Boolean read FBubble3D;
     property BubbleScale: TCT_BubbleScale read FBubbleScale;
     property ShowNegBubbles: TCT_Boolean read FShowNegBubbles;
     property SizeRepresents: TCT_SizeRepresents read FSizeRepresents;
     property AxIdXpgList: TCT_UnsignedIntXpgList read FAxIdXpgList;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_BubbleChartXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_BubbleChart;
public
     function  Add: TCT_BubbleChart;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_BubbleChartXpgList);
     procedure CopyTo(AItem: TCT_BubbleChartXpgList);
     property Items[Index: integer]: TCT_BubbleChart read GetItems; default;
     end;

     TCT_ValAx = class(TXPGBase)
protected
     FEG_AxShared: TEG_AxShared;
     FCrossBetween: TCT_CrossBetween;
     FMajorUnit: TCT_AxisUnit;
     FMinorUnit: TCT_AxisUnit;
     FDispUnits: TCT_DispUnits;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_ValAx);
     procedure CopyTo(AItem: TCT_ValAx);
     function  Create_CrossBetween: TCT_CrossBetween;
     function  Create_MajorUnit: TCT_AxisUnit;
     function  Create_MinorUnit: TCT_AxisUnit;
     function  Create_DispUnits: TCT_DispUnits;
     function  Create_ExtLst: TCT_ExtensionList;

     property EG_AxShared: TEG_AxShared read FEG_AxShared;
     property CrossBetween: TCT_CrossBetween read FCrossBetween;
     property MajorUnit: TCT_AxisUnit read FMajorUnit;
     property MinorUnit: TCT_AxisUnit read FMinorUnit;
     property DispUnits: TCT_DispUnits read FDispUnits;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_ValAxXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_ValAx;
public
     function  Add: TCT_ValAx;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_ValAxXpgList);
     procedure CopyTo(AItem: TCT_ValAxXpgList);
     property Items[Index: integer]: TCT_ValAx read GetItems; default;
     end;

     TCT_CatAx = class(TXPGBase)
protected
     FEG_AxShared: TEG_AxShared;
     FAuto: TCT_Boolean;
     FLblAlgn: TCT_LblAlgn;
     FLblOffset: TCT_LblOffset;
     FTickLblSkip: TCT_Skip;
     FTickMarkSkip: TCT_Skip;
     FNoMultiLvlLbl: TCT_Boolean;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_CatAx);
     procedure CopyTo(AItem: TCT_CatAx);
     function  Create_Auto: TCT_Boolean;
     function  Create_LblAlgn: TCT_LblAlgn;
     function  Create_LblOffset: TCT_LblOffset;
     function  Create_TickLblSkip: TCT_Skip;
     function  Create_TickMarkSkip: TCT_Skip;
     function  Create_NoMultiLvlLbl: TCT_Boolean;
     function  Create_ExtLst: TCT_ExtensionList;

     property EG_AxShared: TEG_AxShared read FEG_AxShared;
     property Auto: TCT_Boolean read FAuto;
     property LblAlgn: TCT_LblAlgn read FLblAlgn;
     property LblOffset: TCT_LblOffset read FLblOffset;
     property TickLblSkip: TCT_Skip read FTickLblSkip;
     property TickMarkSkip: TCT_Skip read FTickMarkSkip;
     property NoMultiLvlLbl: TCT_Boolean read FNoMultiLvlLbl;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_CatAxXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_CatAx;
public
     function  Add: TCT_CatAx;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_CatAxXpgList);
     procedure CopyTo(AItem: TCT_CatAxXpgList);
     property Items[Index: integer]: TCT_CatAx read GetItems; default;
     end;

     TCT_DateAx = class(TXPGBase)
protected
     FEG_AxShared: TEG_AxShared;
     FAuto: TCT_Boolean;
     FLblOffset: TCT_LblOffset;
     FBaseTimeUnit: TCT_TimeUnit;
     FMajorUnit: TCT_AxisUnit;
     FMajorTimeUnit: TCT_TimeUnit;
     FMinorUnit: TCT_AxisUnit;
     FMinorTimeUnit: TCT_TimeUnit;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_DateAx);
     procedure CopyTo(AItem: TCT_DateAx);
     function  Create_Auto: TCT_Boolean;
     function  Create_LblOffset: TCT_LblOffset;
     function  Create_BaseTimeUnit: TCT_TimeUnit;
     function  Create_MajorUnit: TCT_AxisUnit;
     function  Create_MajorTimeUnit: TCT_TimeUnit;
     function  Create_MinorUnit: TCT_AxisUnit;
     function  Create_MinorTimeUnit: TCT_TimeUnit;
     function  Create_ExtLst: TCT_ExtensionList;

     property EG_AxShared: TEG_AxShared read FEG_AxShared;
     property Auto: TCT_Boolean read FAuto;
     property LblOffset: TCT_LblOffset read FLblOffset;
     property BaseTimeUnit: TCT_TimeUnit read FBaseTimeUnit;
     property MajorUnit: TCT_AxisUnit read FMajorUnit;
     property MajorTimeUnit: TCT_TimeUnit read FMajorTimeUnit;
     property MinorUnit: TCT_AxisUnit read FMinorUnit;
     property MinorTimeUnit: TCT_TimeUnit read FMinorTimeUnit;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_DateAxXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_DateAx;
public
     function  Add: TCT_DateAx;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_DateAxXpgList);
     procedure CopyTo(AItem: TCT_DateAxXpgList);
     property Items[Index: integer]: TCT_DateAx read GetItems; default;
     end;

     TCT_SerAx = class(TXPGBase)
protected
     FEG_AxShared: TEG_AxShared;
     FTickLblSkip: TCT_Skip;
     FTickMarkSkip: TCT_Skip;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_SerAx);
     procedure CopyTo(AItem: TCT_SerAx);
     function  Create_TickLblSkip: TCT_Skip;
     function  Create_TickMarkSkip: TCT_Skip;
     function  Create_ExtLst: TCT_ExtensionList;

     property EG_AxShared: TEG_AxShared read FEG_AxShared;
     property TickLblSkip: TCT_Skip read FTickLblSkip;
     property TickMarkSkip: TCT_Skip read FTickMarkSkip;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_SerAxXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_SerAx;
public
     function  Add: TCT_SerAx;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_SerAxXpgList);
     procedure CopyTo(AItem: TCT_SerAxXpgList);
     property Items[Index: integer]: TCT_SerAx read GetItems; default;
     end;

     TCT_DTable = class(TXPGBase)
protected
     FShowHorzBorder: TCT_Boolean;
     FShowVertBorder: TCT_Boolean;
     FShowOutline: TCT_Boolean;
     FShowKeys: TCT_Boolean;
     FSpPr: TCT_ShapeProperties;
     FTxPr: TCT_TextBody;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_DTable);
     procedure CopyTo(AItem: TCT_DTable);
     function  Create_ShowHorzBorder: TCT_Boolean;
     function  Create_ShowVertBorder: TCT_Boolean;
     function  Create_ShowOutline: TCT_Boolean;
     function  Create_ShowKeys: TCT_Boolean;
     function  Create_SpPr: TCT_ShapeProperties;
     function  Create_TxPr: TCT_TextBody;
     function  Create_ExtLst: TCT_ExtensionList;

     property ShowHorzBorder: TCT_Boolean read FShowHorzBorder;
     property ShowVertBorder: TCT_Boolean read FShowVertBorder;
     property ShowOutline: TCT_Boolean read FShowOutline;
     property ShowKeys: TCT_Boolean read FShowKeys;
     property SpPr: TCT_ShapeProperties read FSpPr;
     property TxPr: TCT_TextBody read FTxPr;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_LegendPos = class(TXPGBase)
protected
     FVal: TST_LegendPos;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_LegendPos);
     procedure CopyTo(AItem: TCT_LegendPos);

     property Val: TST_LegendPos read FVal write FVal;
     end;

     TCT_LegendEntry = class(TXPGBase)
protected
     FIdx: TCT_UnsignedInt;
     FDelete: TCT_Boolean;
     FEG_LegendEntryData: TEG_LegendEntryData;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_LegendEntry);
     procedure CopyTo(AItem: TCT_LegendEntry);
     function  Create_Idx: TCT_UnsignedInt;
     function  Create_Delete: TCT_Boolean;
     function  Create_ExtLst: TCT_ExtensionList;

     property Idx: TCT_UnsignedInt read FIdx;
     property Delete: TCT_Boolean read FDelete;
     property EG_LegendEntryData: TEG_LegendEntryData read FEG_LegendEntryData;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_LegendEntryXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_LegendEntry;
public
     function  Add: TCT_LegendEntry;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_LegendEntryXpgList);
     procedure CopyTo(AItem: TCT_LegendEntryXpgList);
     property Items[Index: integer]: TCT_LegendEntry read GetItems; default;
     end;

     TCT_DefaultShapeDefinition = class(TXPGBase)
protected
     FA_SpPr: TCT_ShapeProperties;
     FA_BodyPr: TCT_TextBodyProperties;
     FA_LstStyle: TCT_TextListStyle;
     FA_Style: TCT_ShapeStyle;
     FA_ExtLst: TCT_OfficeArtExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_DefaultShapeDefinition);
     procedure CopyTo(AItem: TCT_DefaultShapeDefinition);
     function  Create_A_SpPr: TCT_ShapeProperties;
     function  Create_A_BodyPr: TCT_TextBodyProperties;
     function  Create_A_LstStyle: TCT_TextListStyle;
     function  Create_A_Style: TCT_ShapeStyle;
     function  Create_A_ExtLst: TCT_OfficeArtExtensionList;

     property A_SpPr: TCT_ShapeProperties read FA_SpPr;
     property A_BodyPr: TCT_TextBodyProperties read FA_BodyPr;
     property A_LstStyle: TCT_TextListStyle read FA_LstStyle;
     property A_Style: TCT_ShapeStyle read FA_Style;
     property A_ExtLst: TCT_OfficeArtExtensionList read FA_ExtLst;
     end;

     TCT_ColorSchemeAndMapping = class(TXPGBase)
protected
     FA_ClrScheme: TCT_ColorScheme;
     FA_ClrMap: TCT_ColorMapping;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_ColorSchemeAndMapping);
     procedure CopyTo(AItem: TCT_ColorSchemeAndMapping);
     function  Create_A_ClrScheme: TCT_ColorScheme;
     function  Create_A_ClrMap: TCT_ColorMapping;

     property A_ClrScheme: TCT_ColorScheme read FA_ClrScheme;
     property A_ClrMap: TCT_ColorMapping read FA_ClrMap;
     end;

     TCT_ColorSchemeAndMappingXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_ColorSchemeAndMapping;
public
     function  Add: TCT_ColorSchemeAndMapping;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_ColorSchemeAndMappingXpgList);
     procedure CopyTo(AItem: TCT_ColorSchemeAndMappingXpgList);
     property Items[Index: integer]: TCT_ColorSchemeAndMapping read GetItems; default;
     end;

     TCT_PivotFmts = class(TXPGBase)
protected
     FPivotFmtXpgList: TCT_PivotFmtXpgList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_PivotFmts);
     procedure CopyTo(AItem: TCT_PivotFmts);
     function  Create_PivotFmtXpgList: TCT_PivotFmtXpgList;

     property PivotFmtXpgList: TCT_PivotFmtXpgList read FPivotFmtXpgList;
     end;

     TCT_View3D = class(TXPGBase)
protected
     FRotX: TCT_RotX;
     FHPercent: TCT_HPercent;
     FRotY: TCT_RotY;
     FDepthPercent: TCT_DepthPercent;
     FRAngAx: TCT_Boolean;
     FPerspective: TCT_Perspective;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_View3D);
     procedure CopyTo(AItem: TCT_View3D);
     function  Create_RotX: TCT_RotX;
     function  Create_HPercent: TCT_HPercent;
     function  Create_RotY: TCT_RotY;
     function  Create_DepthPercent: TCT_DepthPercent;
     function  Create_RAngAx: TCT_Boolean;
     function  Create_Perspective: TCT_Perspective;
     function  Create_ExtLst: TCT_ExtensionList;

     property RotX: TCT_RotX read FRotX;
     property HPercent: TCT_HPercent read FHPercent;
     property RotY: TCT_RotY read FRotY;
     property DepthPercent: TCT_DepthPercent read FDepthPercent;
     property RAngAx: TCT_Boolean read FRAngAx;
     property Perspective: TCT_Perspective read FPerspective;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_Surface = class(TXPGBase)
protected
     FThickness: TCT_UnsignedInt;
     FSpPr: TCT_ShapeProperties;
     FPictureOptions: TCT_PictureOptions;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Surface);
     procedure CopyTo(AItem: TCT_Surface);
     function  Create_Thickness: TCT_UnsignedInt;
     function  Create_SpPr: TCT_ShapeProperties;
     function  Create_PictureOptions: TCT_PictureOptions;
     function  Create_ExtLst: TCT_ExtensionList;

     property Thickness: TCT_UnsignedInt read FThickness;
     property SpPr: TCT_ShapeProperties read FSpPr;
     property PictureOptions: TCT_PictureOptions read FPictureOptions;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_PlotArea = class(TXPGBase)
protected
     FLayout: TCT_Layout;
     FAreaChartXpgList: TCT_AreaChartXpgList;
     FArea3DChartXpgList: TCT_Area3DChartXpgList;
     FLineChartXpgList: TCT_LineChartXpgList;
     FLine3DChartXpgList: TCT_Line3DChartXpgList;
     FStockChartXpgList: TCT_StockChartXpgList;
     FRadarChartXpgList: TCT_RadarChartXpgList;
     FScatterChartXpgList: TCT_ScatterChartXpgList;
     FPieChartXpgList: TCT_PieChartXpgList;
     FPie3DChartXpgList: TCT_Pie3DChartXpgList;
     FDoughnutChartXpgList: TCT_DoughnutChartXpgList;
     FBarChartXpgList: TCT_BarChartXpgList;
     FBar3DChartXpgList: TCT_Bar3DChartXpgList;
     FOfPieChartXpgList: TCT_OfPieChartXpgList;
     FSurfaceChartXpgList: TCT_SurfaceChartXpgList;
     FSurface3DChartXpgList: TCT_Surface3DChartXpgList;
     FBubbleChartXpgList: TCT_BubbleChartXpgList;
     FValAxXpgList: TCT_ValAxXpgList;
     FCatAxXpgList: TCT_CatAxXpgList;
     FDateAxXpgList: TCT_DateAxXpgList;
     FSerAxXpgList: TCT_SerAxXpgList;
     FDTable: TCT_DTable;
     FSpPr: TCT_ShapeProperties;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_PlotArea);
     procedure CopyTo(AItem: TCT_PlotArea);
     function  Create_Layout: TCT_Layout;
     function  Create_AreaChartXpgList: TCT_AreaChartXpgList;
     function  Create_Area3DChartXpgList: TCT_Area3DChartXpgList;
     function  Create_LineChartXpgList: TCT_LineChartXpgList;
     function  Create_Line3DChartXpgList: TCT_Line3DChartXpgList;
     function  Create_StockChartXpgList: TCT_StockChartXpgList;
     function  Create_RadarChartXpgList: TCT_RadarChartXpgList;
     function  Create_ScatterChartXpgList: TCT_ScatterChartXpgList;
     function  Create_PieChartXpgList: TCT_PieChartXpgList;
     function  Create_Pie3DChartXpgList: TCT_Pie3DChartXpgList;
     function  Create_DoughnutChartXpgList: TCT_DoughnutChartXpgList;
     function  Create_BarChartXpgList: TCT_BarChartXpgList;
     function  Create_Bar3DChartXpgList: TCT_Bar3DChartXpgList;
     function  Create_OfPieChartXpgList: TCT_OfPieChartXpgList;
     function  Create_SurfaceChartXpgList: TCT_SurfaceChartXpgList;
     function  Create_Surface3DChartXpgList: TCT_Surface3DChartXpgList;
     function  Create_BubbleChartXpgList: TCT_BubbleChartXpgList;
     function  Create_ValAxXpgList: TCT_ValAxXpgList;
     function  Create_CatAxXpgList: TCT_CatAxXpgList;
     function  Create_DateAxXpgList: TCT_DateAxXpgList;
     function  Create_SerAxXpgList: TCT_SerAxXpgList;
     function  Create_DTable: TCT_DTable;
     function  Create_SpPr: TCT_ShapeProperties;
     function  Create_ExtLst: TCT_ExtensionList;

     property Layout: TCT_Layout read FLayout;
     property AreaChartXpgList: TCT_AreaChartXpgList read FAreaChartXpgList;
     property Area3DChartXpgList: TCT_Area3DChartXpgList read FArea3DChartXpgList;
     property LineChartXpgList: TCT_LineChartXpgList read FLineChartXpgList;
     property Line3DChartXpgList: TCT_Line3DChartXpgList read FLine3DChartXpgList;
     property StockChartXpgList: TCT_StockChartXpgList read FStockChartXpgList;
     property RadarChartXpgList: TCT_RadarChartXpgList read FRadarChartXpgList;
     property ScatterChartXpgList: TCT_ScatterChartXpgList read FScatterChartXpgList;
     property PieChartXpgList: TCT_PieChartXpgList read FPieChartXpgList;
     property Pie3DChartXpgList: TCT_Pie3DChartXpgList read FPie3DChartXpgList;
     property DoughnutChartXpgList: TCT_DoughnutChartXpgList read FDoughnutChartXpgList;
     property BarChartXpgList: TCT_BarChartXpgList read FBarChartXpgList;
     property Bar3DChartXpgList: TCT_Bar3DChartXpgList read FBar3DChartXpgList;
     property OfPieChartXpgList: TCT_OfPieChartXpgList read FOfPieChartXpgList;
     property SurfaceChartXpgList: TCT_SurfaceChartXpgList read FSurfaceChartXpgList;
     property Surface3DChartXpgList: TCT_Surface3DChartXpgList read FSurface3DChartXpgList;
     property BubbleChartXpgList: TCT_BubbleChartXpgList read FBubbleChartXpgList;
     property ValAxXpgList: TCT_ValAxXpgList read FValAxXpgList;
     property CatAxXpgList: TCT_CatAxXpgList read FCatAxXpgList;
     property DateAxXpgList: TCT_DateAxXpgList read FDateAxXpgList;
     property SerAxXpgList: TCT_SerAxXpgList read FSerAxXpgList;
     property DTable: TCT_DTable read FDTable;
     property SpPr: TCT_ShapeProperties read FSpPr;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_Legend = class(TXPGBase)
protected
     FLegendPos: TCT_LegendPos;
     FLegendEntryXpgList: TCT_LegendEntryXpgList;
     FLayout: TCT_Layout;
     FOverlay: TCT_Boolean;
     FSpPr: TCT_ShapeProperties;
     FTxPr: TCT_TextBody;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Legend);
     procedure CopyTo(AItem: TCT_Legend);
     function  Create_LegendPos: TCT_LegendPos;
     function  Create_LegendEntryXpgList: TCT_LegendEntryXpgList;
     function  Create_Layout: TCT_Layout;
     function  Create_Overlay: TCT_Boolean;
     function  Create_SpPr: TCT_ShapeProperties;
     function  Create_TxPr: TCT_TextBody;
     function  Create_ExtLst: TCT_ExtensionList;

     property LegendPos: TCT_LegendPos read FLegendPos;
     property LegendEntryXpgList: TCT_LegendEntryXpgList read FLegendEntryXpgList;
     property Layout: TCT_Layout read FLayout;
     property Overlay: TCT_Boolean read FOverlay;
     property SpPr: TCT_ShapeProperties read FSpPr;
     property TxPr: TCT_TextBody read FTxPr;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_DispBlanksAs = class(TXPGBase)
protected
     FVal: TST_DispBlanksAs;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_DispBlanksAs);
     procedure CopyTo(AItem: TCT_DispBlanksAs);

     property Val: TST_DispBlanksAs read FVal write FVal;
     end;

     TCT_HeaderFooter = class(TXPGBase)
protected
     FAlignWithMargins: boolean;
     FDifferentOddEven: boolean;
     FDifferentFirst: boolean;
     FOddHeader: AxUCString;
     FOddFooter: AxUCString;
     FEvenHeader: AxUCString;
     FEvenFooter: AxUCString;
     FFirstHeader: AxUCString;
     FFirstFooter: AxUCString;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_HeaderFooter);
     procedure CopyTo(AItem: TCT_HeaderFooter);

     property AlignWithMargins: boolean read FAlignWithMargins write FAlignWithMargins;
     property DifferentOddEven: boolean read FDifferentOddEven write FDifferentOddEven;
     property DifferentFirst: boolean read FDifferentFirst write FDifferentFirst;
     property OddHeader: AxUCString read FOddHeader write FOddHeader;
     property OddFooter: AxUCString read FOddFooter write FOddFooter;
     property EvenHeader: AxUCString read FEvenHeader write FEvenHeader;
     property EvenFooter: AxUCString read FEvenFooter write FEvenFooter;
     property FirstHeader: AxUCString read FFirstHeader write FFirstHeader;
     property FirstFooter: AxUCString read FFirstFooter write FFirstFooter;
     end;

     TCT_PageMargins = class(TXPGBase)
protected
     FL: double;
     FR: double;
     FT: double;
     FB: double;
     FHeader: double;
     FFooter: double;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_PageMargins);
     procedure CopyTo(AItem: TCT_PageMargins);

     property L: double read FL write FL;
     property R: double read FR write FR;
     property T: double read FT write FT;
     property B: double read FB write FB;
     property Header: double read FHeader write FHeader;
     property Footer: double read FFooter write FFooter;
     end;

     TCT_PageSetup = class(TXPGBase)
protected
     FPaperSize: integer;
     FFirstPageNumber: integer;
     FOrientation: TST_PageSetupOrientation;
     FBlackAndWhite: boolean;
     FDraft: boolean;
     FUseFirstPageNumber: boolean;
     FHorizontalDpi: integer;
     FVerticalDpi: integer;
     FCopies: integer;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_PageSetup);
     procedure CopyTo(AItem: TCT_PageSetup);

     property PaperSize: integer read FPaperSize write FPaperSize;
     property FirstPageNumber: integer read FFirstPageNumber write FFirstPageNumber;
     property Orientation: TST_PageSetupOrientation read FOrientation write FOrientation;
     property BlackAndWhite: boolean read FBlackAndWhite write FBlackAndWhite;
     property Draft: boolean read FDraft write FDraft;
     property UseFirstPageNumber: boolean read FUseFirstPageNumber write FUseFirstPageNumber;
     property HorizontalDpi: integer read FHorizontalDpi write FHorizontalDpi;
     property VerticalDpi: integer read FVerticalDpi write FVerticalDpi;
     property Copies: integer read FCopies write FCopies;
     end;

     TCT_RelId = class(TXPGBase)
protected
     FR_Id: AxUCString;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_RelId);
     procedure CopyTo(AItem: TCT_RelId);

     property R_Id: AxUCString read FR_Id write FR_Id;
     end;

     TCT_ObjectStyleDefaults = class(TXPGBase)
protected
     FA_SpDef: TCT_DefaultShapeDefinition;
     FA_LnDef: TCT_DefaultShapeDefinition;
     FA_TxDef: TCT_DefaultShapeDefinition;
     FA_ExtLst: TCT_OfficeArtExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_ObjectStyleDefaults);
     procedure CopyTo(AItem: TCT_ObjectStyleDefaults);
     function  Create_A_SpDef: TCT_DefaultShapeDefinition;
     function  Create_A_LnDef: TCT_DefaultShapeDefinition;
     function  Create_A_TxDef: TCT_DefaultShapeDefinition;
     function  Create_A_ExtLst: TCT_OfficeArtExtensionList;

     property A_SpDef: TCT_DefaultShapeDefinition read FA_SpDef;
     property A_LnDef: TCT_DefaultShapeDefinition read FA_LnDef;
     property A_TxDef: TCT_DefaultShapeDefinition read FA_TxDef;
     property A_ExtLst: TCT_OfficeArtExtensionList read FA_ExtLst;
     end;

     TCT_ColorSchemeList = class(TXPGBase)
protected
     FA_ExtraClrSchemeXpgList: TCT_ColorSchemeAndMappingXpgList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_ColorSchemeList);
     procedure CopyTo(AItem: TCT_ColorSchemeList);
     function  Create_A_ExtraClrSchemeXpgList: TCT_ColorSchemeAndMappingXpgList;

     property A_ExtraClrSchemeXpgList: TCT_ColorSchemeAndMappingXpgList read FA_ExtraClrSchemeXpgList;
     end;

     TCT_TextLanguageID = class(TXPGBase)
protected
     FVal: AxUCString;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_TextLanguageID);
     procedure CopyTo(AItem: TCT_TextLanguageID);

     property Val: AxUCString read FVal write FVal;
     end;

     TCT_Style = class(TXPGBase)
protected
     FVal: integer;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Style);
     procedure CopyTo(AItem: TCT_Style);

     property Val: integer read FVal write FVal;
     end;

     TCT_PivotSource = class(TXPGBase)
protected
     FName: AxUCString;
     FFmtId: TCT_UnsignedInt;
     FExtLstXpgList: TCT_ExtensionListXpgList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_PivotSource);
     procedure CopyTo(AItem: TCT_PivotSource);
     function  Create_FmtId: TCT_UnsignedInt;
     function  Create_ExtLstXpgList: TCT_ExtensionListXpgList;

     property Name: AxUCString read FName write FName;
     property FmtId: TCT_UnsignedInt read FFmtId;
     property ExtLstXpgList: TCT_ExtensionListXpgList read FExtLstXpgList;
     end;

     TCT_Protection = class(TXPGBase)
protected
     FChartObject: TCT_Boolean;
     FData: TCT_Boolean;
     FFormatting: TCT_Boolean;
     FSelection: TCT_Boolean;
     FUserInterface: TCT_Boolean;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Protection);
     procedure CopyTo(AItem: TCT_Protection);
     function  Create_ChartObject: TCT_Boolean;
     function  Create_Data: TCT_Boolean;
     function  Create_Formatting: TCT_Boolean;
     function  Create_Selection: TCT_Boolean;
     function  Create_UserInterface: TCT_Boolean;

     property ChartObject: TCT_Boolean read FChartObject;
     property Data: TCT_Boolean read FData;
     property Formatting: TCT_Boolean read FFormatting;
     property Selection: TCT_Boolean read FSelection;
     property UserInterface: TCT_Boolean read FUserInterface;
     end;

     TCT_Chart = class(TXPGBase)
protected
     FTitle: TCT_Title;
     FAutoTitleDeleted: TCT_Boolean;
     FPivotFmts: TCT_PivotFmts;
     FView3D: TCT_View3D;
     FFloor: TCT_Surface;
     FSideWall: TCT_Surface;
     FBackWall: TCT_Surface;
     FPlotArea: TCT_PlotArea;
     FLegend: TCT_Legend;
     FPlotVisOnly: TCT_Boolean;
     FDispBlanksAs: TCT_DispBlanksAs;
     FShowDLblsOverMax: TCT_Boolean;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Chart);
     procedure CopyTo(AItem: TCT_Chart);
     function  Create_Title: TCT_Title;
     function  Create_AutoTitleDeleted: TCT_Boolean;
     function  Create_PivotFmts: TCT_PivotFmts;
     function  Create_View3D: TCT_View3D;
     function  Create_Floor: TCT_Surface;
     function  Create_SideWall: TCT_Surface;
     function  Create_BackWall: TCT_Surface;
     function  Create_PlotArea: TCT_PlotArea;
     function  Create_Legend: TCT_Legend;
     function  Create_PlotVisOnly: TCT_Boolean;
     function  Create_DispBlanksAs: TCT_DispBlanksAs;
     function  Create_ShowDLblsOverMax: TCT_Boolean;
     function  Create_ExtLst: TCT_ExtensionList;

     property Title: TCT_Title read FTitle;
     property AutoTitleDeleted: TCT_Boolean read FAutoTitleDeleted;
     property PivotFmts: TCT_PivotFmts read FPivotFmts;
     property View3D: TCT_View3D read FView3D;
     property Floor: TCT_Surface read FFloor;
     property SideWall: TCT_Surface read FSideWall;
     property BackWall: TCT_Surface read FBackWall;
     property PlotArea: TCT_PlotArea read FPlotArea;
     property Legend: TCT_Legend read FLegend;
     property PlotVisOnly: TCT_Boolean read FPlotVisOnly;
     property DispBlanksAs: TCT_DispBlanksAs read FDispBlanksAs;
     property ShowDLblsOverMax: TCT_Boolean read FShowDLblsOverMax;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_ExternalData = class(TXPGBase)
protected
     FR_Id: AxUCString;
     FAutoUpdate: TCT_Boolean;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_ExternalData);
     procedure CopyTo(AItem: TCT_ExternalData);
     function  Create_AutoUpdate: TCT_Boolean;

     property R_Id: AxUCString read FR_Id write FR_Id;
     property AutoUpdate: TCT_Boolean read FAutoUpdate;
     end;

     TCT_PrintSettings = class(TXPGBase)
protected
     FHeaderFooter: TCT_HeaderFooter;
     FPageMargins: TCT_PageMargins;
     FPageSetup: TCT_PageSetup;
     FLegacyDrawingHF: TCT_RelId;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_PrintSettings);
     procedure CopyTo(AItem: TCT_PrintSettings);
     function  Create_HeaderFooter: TCT_HeaderFooter;
     function  Create_PageMargins: TCT_PageMargins;
     function  Create_PageSetup: TCT_PageSetup;
     function  Create_LegacyDrawingHF: TCT_RelId;

     property HeaderFooter: TCT_HeaderFooter read FHeaderFooter;
     property PageMargins: TCT_PageMargins read FPageMargins;
     property PageSetup: TCT_PageSetup read FPageSetup;
     property LegacyDrawingHF: TCT_RelId read FLegacyDrawingHF;
     end;

     TCT_EmptyElement = class(TXPGBase)
public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_EmptyElement);
     procedure CopyTo(AItem: TCT_EmptyElement);

     end;

     TCT_BaseStylesOverride = class(TXPGBase)
protected
     FA_ClrScheme: TCT_ColorScheme;
     FA_FontScheme: TCT_FontScheme;
     FA_FmtScheme: TCT_StyleMatrix;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_BaseStylesOverride);
     procedure CopyTo(AItem: TCT_BaseStylesOverride);
     function  Create_A_ClrScheme: TCT_ColorScheme;
     function  Create_A_FontScheme: TCT_FontScheme;
     function  Create_A_FmtScheme: TCT_StyleMatrix;

     property A_ClrScheme: TCT_ColorScheme read FA_ClrScheme;
     property A_FontScheme: TCT_FontScheme read FA_FontScheme;
     property A_FmtScheme: TCT_StyleMatrix read FA_FmtScheme;
     end;

     TCT_OfficeStyleSheet = class(TXPGBase)
protected
     FName: AxUCString;
     FA_ThemeElements: TCT_BaseStyles;
     FA_ObjectDefaults: TCT_ObjectStyleDefaults;
     FA_ExtraClrSchemeLst: TCT_ColorSchemeList;
     FA_CustClrLst: TCT_CustomColorList;
     FA_ExtLst: TCT_OfficeArtExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_OfficeStyleSheet);
     procedure CopyTo(AItem: TCT_OfficeStyleSheet);
     function  Create_A_ThemeElements: TCT_BaseStyles;
     function  Create_A_ObjectDefaults: TCT_ObjectStyleDefaults;
     function  Create_A_ExtraClrSchemeLst: TCT_ColorSchemeList;
     function  Create_A_CustClrLst: TCT_CustomColorList;
     function  Create_A_ExtLst: TCT_OfficeArtExtensionList;

     property Name: AxUCString read FName write FName;
     property A_ThemeElements: TCT_BaseStyles read FA_ThemeElements;
     property A_ObjectDefaults: TCT_ObjectStyleDefaults read FA_ObjectDefaults;
     property A_ExtraClrSchemeLst: TCT_ColorSchemeList read FA_ExtraClrSchemeLst;
     property A_CustClrLst: TCT_CustomColorList read FA_CustClrLst;
     property A_ExtLst: TCT_OfficeArtExtensionList read FA_ExtLst;
     end;

     TCT_ChartSpace = class(TXPGBase)
protected
     FDate1904: TCT_Boolean;
     FLang: TCT_TextLanguageID;
     FRoundedCorners: TCT_Boolean;
     FStyle: TCT_Style;
     FClrMapOvr: TCT_ColorMapping;
     FPivotSource: TCT_PivotSource;
     FProtection: TCT_Protection;
     FChart: TCT_Chart;
     FSpPr: TCT_ShapeProperties;
     FTxPr: TCT_TextBody;
     FExternalData: TCT_ExternalData;
     FPrintSettings: TCT_PrintSettings;
     FUserShapes: TCT_RelId;
     FExtLst: TCT_ExtensionList;
public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_ChartSpace);
     procedure CopyTo(AItem: TCT_ChartSpace);
     function  Create_Date1904: TCT_Boolean;
     function  Create_Lang: TCT_TextLanguageID;
     function  Create_RoundedCorners: TCT_Boolean;
     function  Create_Style: TCT_Style;
     function  Create_ClrMapOvr: TCT_ColorMapping;
     function  Create_PivotSource: TCT_PivotSource;
     function  Create_Protection: TCT_Protection;
     function  Create_Chart: TCT_Chart;
     function  Create_SpPr: TCT_ShapeProperties;
     function  Create_TxPr: TCT_TextBody;
     function  Create_ExternalData: TCT_ExternalData;
     function  Create_PrintSettings: TCT_PrintSettings;
     function  Create_UserShapes: TCT_RelId;
     function  Create_ExtLst: TCT_ExtensionList;

     property Date1904: TCT_Boolean read FDate1904;
     property Lang: TCT_TextLanguageID read FLang;
     property RoundedCorners: TCT_Boolean read FRoundedCorners;
     property Style: TCT_Style read FStyle;
     property ClrMapOvr: TCT_ColorMapping read FClrMapOvr;
     property PivotSource: TCT_PivotSource read FPivotSource;
     property Protection: TCT_Protection read FProtection;
     property Chart: TCT_Chart read FChart;
     property SpPr: TCT_ShapeProperties read FSpPr;
     property TxPr: TCT_TextBody read FTxPr;
     property ExternalData: TCT_ExternalData read FExternalData;
     property PrintSettings: TCT_PrintSettings read FPrintSettings;
     property UserShapes: TCT_RelId read FUserShapes;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_ColorMappingOverride = class(TXPGBase)
protected
     FA_MasterClrMapping: TCT_EmptyElement;
     FA_OverrideClrMapping: TCT_ColorMapping;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_ColorMappingOverride);
     procedure CopyTo(AItem: TCT_ColorMappingOverride);
     function  Create_A_MasterClrMapping: TCT_EmptyElement;
     function  Create_A_OverrideClrMapping: TCT_ColorMapping;

     property A_MasterClrMapping: TCT_EmptyElement read FA_MasterClrMapping;
     property A_OverrideClrMapping: TCT_ColorMapping read FA_OverrideClrMapping;
     end;

     TCT_ClipboardStyleSheet = class(TXPGBase)
protected
     FA_ThemeElements: TCT_BaseStyles;
     FA_ClrMap: TCT_ColorMapping;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_ClipboardStyleSheet);
     procedure CopyTo(AItem: TCT_ClipboardStyleSheet);
     function  Create_A_ThemeElements: TCT_BaseStyles;
     function  Create_A_ClrMap: TCT_ColorMapping;

     property A_ThemeElements: TCT_BaseStyles read FA_ThemeElements;
     property A_ClrMap: TCT_ColorMapping read FA_ClrMap;
     end;

     T__ROOT__ = class(TXPGBase)
protected
     FRootAttributes: TStringXpgList;
     FCurrWriteClass: TClass;
     FChart: TCT_RelId;
     FChartSpace: TCT_ChartSpace;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: T__ROOT__);
     procedure CopyTo(AItem: T__ROOT__);
     function  Create_Chart: TCT_RelId;
     function  Create_ChartSpace: TCT_ChartSpace;

     property RootAttributes: TStringXpgList read FRootAttributes;
     property Chart: TCT_RelId read FChart;
     property ChartSpace: TCT_ChartSpace read FChartSpace;
     end;

     TXPGDocXLSXChart = class(TXPGDocBase)
protected
     FRoot    : T__ROOT__;
     FReader  : TXPGReader;
     FWriter  : TXpgWriteXML;

     FStyle   : TXpgSimpleDOM;
     FColors  : TXpgSimpleDOM;

     function  GetChart: TCT_RelId;
     function  GetChartSpace: TCT_ChartSpace;
public
     constructor Create;
     destructor Destroy; override;

     procedure AddStyle;
     procedure AddColors;

     procedure LoadFromFile(AFilename: AxUCString);
     procedure LoadFromStream(AStream: TStream);
     procedure SaveToFile(AFilename: AxUCString; AClassToWrite: TClass);

     procedure SaveToStream(AStream: TStream);

     property Root: T__ROOT__ read FRoot;
     property Chart: TCT_RelId read GetChart;
     property ChartSpace: TCT_ChartSpace read GetChartSpace;

     property Style   : TXpgSimpleDOM read FStyle;
     property Colors  : TXpgSimpleDOM read FColors;
     end;


implementation

procedure ReadUnionTST_AdjAngle(AValue: AxUCString; APtr: PST_AdjAngle);
begin
  if XmlTryStrToInt(AValue,APtr.Val1) then 
  begin
    APtr.nVal := 0;
    Exit;
  end;
  if AValue <> '' then 
  begin
    APtr.Val2 := AValue;
    APtr.nVal := 1;
    Exit;
  end;
end;

function  WriteUnionTST_AdjAngle(APtr: PST_AdjAngle): AxUCString;
begin
  case APtr.nVal of
    0: Result := XmlIntToStr(APtr.Val1);
    1: Result := APtr.Val2;
  end;
end;

procedure ReadUnionTST_AdjCoordinate(AValue: AxUCString; APtr: PST_AdjCoordinate);
begin
  if XmlTryStrToInt(AValue,APtr.Val1) then 
  begin
    APtr.nVal := 0;
    Exit;
  end;
  if AValue <> '' then 
  begin
    APtr.Val2 := AValue;
    APtr.nVal := 1;
    Exit;
  end;
end;

function  WriteUnionTST_AdjCoordinate(APtr: PST_AdjCoordinate): AxUCString;
begin
  case APtr.nVal of
    0: Result := XmlIntToStr(APtr.Val1);
    1: Result := APtr.Val2;
  end;
end;

{ TCT_Extension }

function  TCT_Extension.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FUri <> '' then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FAnyElements.Count);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_Extension.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := FAnyElements.Add(AReader.QName,AReader.Text);
  if Result <> Self then
    Result.Assigneds := [xaRead];
end;

procedure TCT_Extension.Write(AWriter: TXpgWriteXML);
begin
  FAnyElements.Write(AWriter);
end;

procedure TCT_Extension.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FUri <> '' then
    AWriter.AddAttribute('uri',FUri);
  if FXmlns <> '' then
    AWriter.AddAttribute(FXmlns,FXmlnsVal);
end;

procedure TCT_Extension.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do begin
    if AAttributes[i] = 'uri' then
      FUri := AAttributes.Values[i]
    else if Copy(AAttributes[i],1,5) = 'xmlns' then begin
      FXmlns := AAttributes[i];
      FXmlnsVal := AAttributes.Values[i];
    end
    else
      FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
  end;
end;

constructor TCT_Extension.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FAnyElements := TXPGAnyElements.Create;
  FUri := '';
end;

destructor TCT_Extension.Destroy;
begin
  FAnyElements.Free;
end;

procedure TCT_Extension.Clear;
begin
  FAssigneds := [];
  FAnyElements.Clear;
  FUri := '';
end;

procedure TCT_Extension.Assign(AItem: TCT_Extension);
begin
end;

procedure TCT_Extension.CopyTo(AItem: TCT_Extension);
begin
end;

{ TCT_ExtensionXpgList }

function  TCT_ExtensionXpgList.GetItems(Index: integer): TCT_Extension;
begin
  Result := TCT_Extension(inherited Items[Index]);
end;

function  TCT_ExtensionXpgList.Add: TCT_Extension;
begin
  Result := TCT_Extension.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_ExtensionXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_ExtensionXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_ExtensionXpgList.Assign(AItem: TCT_ExtensionXpgList);
begin
end;

procedure TCT_ExtensionXpgList.CopyTo(AItem: TCT_ExtensionXpgList);
begin
end;

{ TCT_UnsignedInt }

function  TCT_UnsignedInt.CheckAssigned: integer;
begin
  FAssigneds := [xaAttributes];
  Result := 1;
end;

procedure TCT_UnsignedInt.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_UnsignedInt.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('val',XmlIntToStr(FVal));
end;

procedure TCT_UnsignedInt.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then 
    FVal := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_UnsignedInt.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := 2147483632;
end;

destructor TCT_UnsignedInt.Destroy;
begin
end;

procedure TCT_UnsignedInt.Clear;
begin
  FAssigneds := [];
  FVal := 2147483632;
end;

procedure TCT_UnsignedInt.Assign(AItem: TCT_UnsignedInt);
begin
end;

procedure TCT_UnsignedInt.CopyTo(AItem: TCT_UnsignedInt);
begin
end;

{ TCT_UnsignedIntXpgList }

function  TCT_UnsignedIntXpgList.GetItems(Index: integer): TCT_UnsignedInt;
begin
  Result := TCT_UnsignedInt(inherited Items[Index]);
end;

function  TCT_UnsignedIntXpgList.Add: TCT_UnsignedInt;
begin
  Result := TCT_UnsignedInt.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_UnsignedIntXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_UnsignedIntXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_UnsignedIntXpgList.Assign(AItem: TCT_UnsignedIntXpgList);
begin
end;

procedure TCT_UnsignedIntXpgList.CopyTo(AItem: TCT_UnsignedIntXpgList);
begin
end;

{ TCT_StrVal }

function  TCT_StrVal.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [xaAttributes];
  if FV <> '' then 
    Inc(ElemsAssigned);
  Result := 1;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_StrVal.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'c:v' then 
    FV := AReader.Text
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_StrVal.Write(AWriter: TXpgWriteXML);
begin
  AWriter.SimpleTextTag('c:v',FV);
end;

procedure TCT_StrVal.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('idx',XmlIntToStr(FIdx));
end;

procedure TCT_StrVal.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'idx' then 
    FIdx := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_StrVal.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FIdx := 2147483632;
end;

destructor TCT_StrVal.Destroy;
begin
end;

procedure TCT_StrVal.Clear;
begin
  FAssigneds := [];
  FV := '';
  FIdx := 2147483632;
end;

procedure TCT_StrVal.Assign(AItem: TCT_StrVal);
begin
end;

procedure TCT_StrVal.CopyTo(AItem: TCT_StrVal);
begin
end;

{ TCT_StrValXpgList }

function  TCT_StrValXpgList.GetItems(Index: integer): TCT_StrVal;
begin
  Result := TCT_StrVal(inherited Items[Index]);
end;

function  TCT_StrValXpgList.Add: TCT_StrVal;
begin
  Result := TCT_StrVal.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_StrValXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_StrValXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_StrValXpgList.Assign(AItem: TCT_StrValXpgList);
begin
end;

procedure TCT_StrValXpgList.CopyTo(AItem: TCT_StrValXpgList);
begin
end;

{ TCT_ExtensionList }

function  TCT_ExtensionList.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FExtXpgList <> Nil then 
    Inc(ElemsAssigned,FExtXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_ExtensionList.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'c:ext' then 
  begin
    if FExtXpgList = Nil then 
      FExtXpgList := TCT_ExtensionXpgList.Create(FOwner);
    Result := FExtXpgList.Add;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_ExtensionList.Write(AWriter: TXpgWriteXML);
begin
  if FExtXpgList <> Nil then 
    FExtXpgList.Write(AWriter,'c:ext');
end;

constructor TCT_ExtensionList.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
end;

destructor TCT_ExtensionList.Destroy;
begin
  if FExtXpgList <> Nil then 
    FExtXpgList.Free;
end;

procedure TCT_ExtensionList.Clear;
begin
  FAssigneds := [];
  if FExtXpgList <> Nil then 
    FreeAndNil(FExtXpgList);
end;

procedure TCT_ExtensionList.Assign(AItem: TCT_ExtensionList);
begin
end;

procedure TCT_ExtensionList.CopyTo(AItem: TCT_ExtensionList);
begin
end;

function  TCT_ExtensionList.Create_ExtXpgList: TCT_ExtensionXpgList;
begin
  Result := TCT_ExtensionXpgList.Create(FOwner);
  FExtXpgList := Result;
end;

{ TCT_ExtensionListXpgList }

function  TCT_ExtensionListXpgList.GetItems(Index: integer): TCT_ExtensionList;
begin
  Result := TCT_ExtensionList(inherited Items[Index]);
end;

function  TCT_ExtensionListXpgList.Add: TCT_ExtensionList;
begin
  Result := TCT_ExtensionList.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_ExtensionListXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_ExtensionListXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
end;

procedure TCT_ExtensionListXpgList.Assign(AItem: TCT_ExtensionListXpgList);
begin
end;

procedure TCT_ExtensionListXpgList.CopyTo(AItem: TCT_ExtensionListXpgList);
begin
end;

{ TCT_LayoutTarget }

function  TCT_LayoutTarget.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FVal <> stltOuter then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_LayoutTarget.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_LayoutTarget.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FVal <> stltOuter then 
    AWriter.AddAttribute('val',StrTST_LayoutTarget[Integer(FVal)]);
end;

procedure TCT_LayoutTarget.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then 
    FVal := TST_LayoutTarget(StrToEnum('stlt' + AAttributes.Values[0]))
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_LayoutTarget.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := stltOuter;
end;

destructor TCT_LayoutTarget.Destroy;
begin
end;

procedure TCT_LayoutTarget.Clear;
begin
  FAssigneds := [];
  FVal := stltOuter;
end;

procedure TCT_LayoutTarget.Assign(AItem: TCT_LayoutTarget);
begin
end;

procedure TCT_LayoutTarget.CopyTo(AItem: TCT_LayoutTarget);
begin
end;

{ TCT_LayoutMode }

function  TCT_LayoutMode.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FVal <> stlmFactor then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_LayoutMode.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_LayoutMode.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FVal <> stlmFactor then 
    AWriter.AddAttribute('val',StrTST_LayoutMode[Integer(FVal)]);
end;

procedure TCT_LayoutMode.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then 
    FVal := TST_LayoutMode(StrToEnum('stlm' + AAttributes.Values[0]))
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_LayoutMode.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := stlmFactor;
end;

destructor TCT_LayoutMode.Destroy;
begin
end;

procedure TCT_LayoutMode.Clear;
begin
  FAssigneds := [];
  FVal := stlmFactor;
end;

procedure TCT_LayoutMode.Assign(AItem: TCT_LayoutMode);
begin
end;

procedure TCT_LayoutMode.CopyTo(AItem: TCT_LayoutMode);
begin
end;

{ TCT_Double }

function  TCT_Double.CheckAssigned: integer;
begin
  FAssigneds := [xaAttributes];
  Result := 1;
end;

procedure TCT_Double.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_Double.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('val',XmlFloatToStr(FVal));
end;

procedure TCT_Double.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then 
    FVal := XmlStrToFloatDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_Double.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := NaN;
end;

destructor TCT_Double.Destroy;
begin
end;

procedure TCT_Double.Clear;
begin
  FAssigneds := [];
  FVal := NaN;
end;

procedure TCT_Double.Assign(AItem: TCT_Double);
begin
end;

procedure TCT_Double.CopyTo(AItem: TCT_Double);
begin
end;

{ TCT_StrData }

function  TCT_StrData.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FPtCount <> Nil then 
    Inc(ElemsAssigned,FPtCount.CheckAssigned);
  if FPtXpgList <> Nil then 
    Inc(ElemsAssigned,FPtXpgList.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_StrData.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $0000038A: begin
      if FPtCount = Nil then 
        FPtCount := TCT_UnsignedInt.Create(FOwner);
      Result := FPtCount;
    end;
    $00000181: begin
      if FPtXpgList = Nil then 
        FPtXpgList := TCT_StrValXpgList.Create(FOwner);
      Result := FPtXpgList.Add;
    end;
    $00000321: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_StrData.Write(AWriter: TXpgWriteXML);
begin
  if (FPtCount <> Nil) and FPtCount.Assigned then 
  begin
    FPtCount.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:ptCount');
  end;
  if FPtXpgList <> Nil then 
    FPtXpgList.Write(AWriter,'c:pt');
{$ifdef XLS_WRITE_EXTLST}
  if (FExtLst <> Nil) and FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then
    begin
      AWriter.BeginTag('c:extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:extLst');
{$endif}
end;

constructor TCT_StrData.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 3;
  FAttributeCount := 0;
end;

destructor TCT_StrData.Destroy;
begin
  if FPtCount <> Nil then 
    FPtCount.Free;
  if FPtXpgList <> Nil then 
    FPtXpgList.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_StrData.Clear;
begin
  FAssigneds := [];
  if FPtCount <> Nil then 
    FreeAndNil(FPtCount);
  if FPtXpgList <> Nil then 
    FreeAndNil(FPtXpgList);
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
end;

procedure TCT_StrData.Assign(AItem: TCT_StrData);
begin
end;

procedure TCT_StrData.CopyTo(AItem: TCT_StrData);
begin
end;

function  TCT_StrData.Create_PtCount: TCT_UnsignedInt;
begin
  Result := TCT_UnsignedInt.Create(FOwner);
  FPtCount := Result;
end;

function  TCT_StrData.Create_PtXpgList: TCT_StrValXpgList;
begin
  Result := TCT_StrValXpgList.Create(FOwner);
  FPtXpgList := Result;
end;

function  TCT_StrData.Create_ExtLst: TCT_ExtensionList;
begin
  Result := TCT_ExtensionList.Create(FOwner);
  FExtLst := Result;
end;

{ TCT_ManualLayout }

function  TCT_ManualLayout.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FLayoutTarget <> Nil then 
    Inc(ElemsAssigned,FLayoutTarget.CheckAssigned);
  if FXMode <> Nil then 
    Inc(ElemsAssigned,FXMode.CheckAssigned);
  if FYMode <> Nil then 
    Inc(ElemsAssigned,FYMode.CheckAssigned);
  if FWMode <> Nil then 
    Inc(ElemsAssigned,FWMode.CheckAssigned);
  if FHMode <> Nil then 
    Inc(ElemsAssigned,FHMode.CheckAssigned);
  if FX <> Nil then 
    Inc(ElemsAssigned,FX.CheckAssigned);
  if FY <> Nil then 
    Inc(ElemsAssigned,FY.CheckAssigned);
  if FW <> Nil then 
    Inc(ElemsAssigned,FW.CheckAssigned);
  if FH <> Nil then 
    Inc(ElemsAssigned,FH.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_ManualLayout.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $000005A2: begin
      if FLayoutTarget = Nil then 
        FLayoutTarget := TCT_LayoutTarget.Create(FOwner);
      Result := FLayoutTarget;
    end;
    $0000029A: begin
      if FXMode = Nil then 
        FXMode := TCT_LayoutMode.Create(FOwner);
      Result := FXMode;
    end;
    $0000029B: begin
      if FYMode = Nil then 
        FYMode := TCT_LayoutMode.Create(FOwner);
      Result := FYMode;
    end;
    $00000299: begin
      if FWMode = Nil then 
        FWMode := TCT_LayoutMode.Create(FOwner);
      Result := FWMode;
    end;
    $0000028A: begin
      if FHMode = Nil then 
        FHMode := TCT_LayoutMode.Create(FOwner);
      Result := FHMode;
    end;
    $00000115: begin
      if FX = Nil then 
        FX := TCT_Double.Create(FOwner);
      Result := FX;
    end;
    $00000116: begin
      if FY = Nil then 
        FY := TCT_Double.Create(FOwner);
      Result := FY;
    end;
    $00000114: begin
      if FW = Nil then 
        FW := TCT_Double.Create(FOwner);
      Result := FW;
    end;
    $00000105: begin
      if FH = Nil then 
        FH := TCT_Double.Create(FOwner);
      Result := FH;
    end;
    $00000321: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_ManualLayout.Write(AWriter: TXpgWriteXML);
begin
  if (FLayoutTarget <> Nil) and FLayoutTarget.Assigned then 
  begin
    FLayoutTarget.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:layoutTarget');
  end;
  if (FXMode <> Nil) and FXMode.Assigned then 
  begin
    FXMode.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:xMode');
  end;
  if (FYMode <> Nil) and FYMode.Assigned then 
  begin
    FYMode.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:yMode');
  end;
  if (FWMode <> Nil) and FWMode.Assigned then 
  begin
    FWMode.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:wMode');
  end;
  if (FHMode <> Nil) and FHMode.Assigned then 
  begin
    FHMode.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:hMode');
  end;
  if (FX <> Nil) and FX.Assigned then 
  begin
    FX.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:x');
  end;
  if (FY <> Nil) and FY.Assigned then 
  begin
    FY.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:y');
  end;
  if (FW <> Nil) and FW.Assigned then 
  begin
    FW.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:w');
  end;
  if (FH <> Nil) and FH.Assigned then 
  begin
    FH.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:h');
  end;
{$ifdef XLS_WRITE_EXTLST}
  if (FExtLst <> Nil) and FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then
    begin
      AWriter.BeginTag('c:extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:extLst');
{$endif}
end;

constructor TCT_ManualLayout.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 10;
  FAttributeCount := 0;
end;

destructor TCT_ManualLayout.Destroy;
begin
  if FLayoutTarget <> Nil then 
    FLayoutTarget.Free;
  if FXMode <> Nil then 
    FXMode.Free;
  if FYMode <> Nil then 
    FYMode.Free;
  if FWMode <> Nil then 
    FWMode.Free;
  if FHMode <> Nil then 
    FHMode.Free;
  if FX <> Nil then 
    FX.Free;
  if FY <> Nil then 
    FY.Free;
  if FW <> Nil then 
    FW.Free;
  if FH <> Nil then 
    FH.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_ManualLayout.Clear;
begin
  FAssigneds := [];
  if FLayoutTarget <> Nil then 
    FreeAndNil(FLayoutTarget);
  if FXMode <> Nil then 
    FreeAndNil(FXMode);
  if FYMode <> Nil then 
    FreeAndNil(FYMode);
  if FWMode <> Nil then 
    FreeAndNil(FWMode);
  if FHMode <> Nil then 
    FreeAndNil(FHMode);
  if FX <> Nil then 
    FreeAndNil(FX);
  if FY <> Nil then 
    FreeAndNil(FY);
  if FW <> Nil then 
    FreeAndNil(FW);
  if FH <> Nil then 
    FreeAndNil(FH);
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
end;

procedure TCT_ManualLayout.Assign(AItem: TCT_ManualLayout);
begin
end;

procedure TCT_ManualLayout.CopyTo(AItem: TCT_ManualLayout);
begin
end;

function  TCT_ManualLayout.Create_LayoutTarget: TCT_LayoutTarget;
begin
  Result := TCT_LayoutTarget.Create(FOwner);
  FLayoutTarget := Result;
end;

function  TCT_ManualLayout.Create_XMode: TCT_LayoutMode;
begin
  Result := TCT_LayoutMode.Create(FOwner);
  FXMode := Result;
end;

function  TCT_ManualLayout.Create_YMode: TCT_LayoutMode;
begin
  Result := TCT_LayoutMode.Create(FOwner);
  FYMode := Result;
end;

function  TCT_ManualLayout.Create_WMode: TCT_LayoutMode;
begin
  Result := TCT_LayoutMode.Create(FOwner);
  FWMode := Result;
end;

function  TCT_ManualLayout.Create_HMode: TCT_LayoutMode;
begin
  Result := TCT_LayoutMode.Create(FOwner);
  FHMode := Result;
end;

function  TCT_ManualLayout.Create_X: TCT_Double;
begin
  Result := TCT_Double.Create(FOwner);
  FX := Result;
end;

function  TCT_ManualLayout.Create_Y: TCT_Double;
begin
  Result := TCT_Double.Create(FOwner);
  FY := Result;
end;

function  TCT_ManualLayout.Create_W: TCT_Double;
begin
  Result := TCT_Double.Create(FOwner);
  FW := Result;
end;

function  TCT_ManualLayout.Create_H: TCT_Double;
begin
  Result := TCT_Double.Create(FOwner);
  FH := Result;
end;

function  TCT_ManualLayout.Create_ExtLst: TCT_ExtensionList;
begin
  Result := TCT_ExtensionList.Create(FOwner);
  FExtLst := Result;
end;

{ TCT_StrRef }

function  TCT_StrRef.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FF <> '' then 
    Inc(ElemsAssigned);
  if FStrCache <> Nil then 
    Inc(ElemsAssigned,FStrCache.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_StrRef.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $00000103: FF := AReader.Text;
    $000003CA: begin
      if FStrCache = Nil then 
        FStrCache := TCT_StrData.Create(FOwner);
      Result := FStrCache;
    end;
    $00000321: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_StrRef.Write(AWriter: TXpgWriteXML);
begin
  if FF <> '' then 
    AWriter.SimpleTextTag('c:f',FF);
  if (FStrCache <> Nil) and FStrCache.Assigned then 
    if xaElements in FStrCache.FAssigneds then 
    begin
      AWriter.BeginTag('c:strCache');
      FStrCache.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:strCache');
{$ifdef XLS_WRITE_EXTLST}
  if (FExtLst <> Nil) and FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then
    begin
      AWriter.BeginTag('c:extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:extLst');
{$endif}
end;

constructor TCT_StrRef.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 3;
  FAttributeCount := 0;
end;

destructor TCT_StrRef.Destroy;
begin
  if FStrCache <> Nil then 
    FStrCache.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_StrRef.Clear;
begin
  FAssigneds := [];
  FF := '';
  if FStrCache <> Nil then 
    FreeAndNil(FStrCache);
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
end;

procedure TCT_StrRef.Assign(AItem: TCT_StrRef);
begin
end;

procedure TCT_StrRef.CopyTo(AItem: TCT_StrRef);
begin
end;

function  TCT_StrRef.Create_StrCache: TCT_StrData;
begin
  Result := TCT_StrData.Create(FOwner);
  FStrCache := Result;
end;

function  TCT_StrRef.Create_ExtLst: TCT_ExtensionList;
begin
  Result := TCT_ExtensionList.Create(FOwner);
  FExtLst := Result;
end;

{ TCT_NumFmt }

function  TCT_NumFmt.CheckAssigned: integer;
begin
  FAssigneds := [xaAttributes];
  Result := 1;
end;

procedure TCT_NumFmt.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_NumFmt.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('formatCode',FFormatCode);
  if Byte(FSourceLinked) <> 2 then 
    AWriter.AddAttribute('sourceLinked',XmlBoolToStr(FSourceLinked));
end;

procedure TCT_NumFmt.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $00000404: FFormatCode := AAttributes.Values[i];
      $000004E8: FSourceLinked := XmlStrToBoolDef(AAttributes.Values[i],False);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_NumFmt.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 2;
  Byte(FSourceLinked) := 2;
end;

destructor TCT_NumFmt.Destroy;
begin
end;

procedure TCT_NumFmt.Clear;
begin
  FAssigneds := [];
  FFormatCode := '';
  Byte(FSourceLinked) := 2;
end;

procedure TCT_NumFmt.Assign(AItem: TCT_NumFmt);
begin
end;

procedure TCT_NumFmt.CopyTo(AItem: TCT_NumFmt);
begin
end;

{ TCT_DLblPos }

function  TCT_DLblPos.CheckAssigned: integer;
begin
  FAssigneds := [xaAttributes];
  if FVal <> TST_DLblPos(XPG_UNKNOWN_ENUM) then
    Result := 1
  else
    Result := 0;
end;

procedure TCT_DLblPos.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_DLblPos.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FVal <> TST_DLblPos(XPG_UNKNOWN_ENUM) then
    AWriter.AddAttribute('val',StrTST_DLblPos[Integer(FVal)]);
end;

procedure TCT_DLblPos.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then 
    FVal := TST_DLblPos(StrToEnum('stdlp' + AAttributes.Values[0]))
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_DLblPos.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  Clear;
end;

destructor TCT_DLblPos.Destroy;
begin
end;

procedure TCT_DLblPos.Clear;
begin
  FAssigneds := [];
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := TST_DLblPos(XPG_UNKNOWN_ENUM);
end;

procedure TCT_DLblPos.Assign(AItem: TCT_DLblPos);
begin
end;

procedure TCT_DLblPos.CopyTo(AItem: TCT_DLblPos);
begin
end;

{ TCT_Boolean }

function  TCT_Boolean.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if Byte(FVal) <> 2 then
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then
    FAssigneds := [xaAttributes];
end;

procedure TCT_Boolean.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_Boolean.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if Byte(FVal) <> 2 then
    AWriter.AddAttribute('val',XmlBoolToStr(FVal));
end;

procedure TCT_Boolean.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then
    FVal := XmlStrToBoolDef(AAttributes.Values[0],True)
  else
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_Boolean.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  Byte(FVal) := 2;
end;

destructor TCT_Boolean.Destroy;
begin
end;

procedure TCT_Boolean.Clear;
begin
  FAssigneds := [];
  FVal := True;
end;

procedure TCT_Boolean.Assign(AItem: TCT_Boolean);
begin
end;

procedure TCT_Boolean.CopyTo(AItem: TCT_Boolean);
begin
end;

{ TCT_NumVal }

function  TCT_NumVal.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  // _TODO_001_
  ElemsAssigned := 0;
  FAssigneds := [xaAttributes];
  if FV <> '' then
    Inc(ElemsAssigned);
  Result := 1;
  if ElemsAssigned > 0 then
    FAssigneds := FAssigneds + [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_NumVal.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'c:v' then 
    FV := AReader.Text
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_NumVal.Write(AWriter: TXpgWriteXML);
begin
  AWriter.SimpleTextTag('c:v',FV);
end;

procedure TCT_NumVal.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('idx',XmlIntToStr(FIdx));
  if FFormatCode <> '' then 
    AWriter.AddAttribute('formatCode',FFormatCode);
end;

procedure TCT_NumVal.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $00000145: FIdx := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000404: FFormatCode := AAttributes.Values[i];
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_NumVal.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 2;
  FIdx := 2147483632;
end;

destructor TCT_NumVal.Destroy;
begin
end;

procedure TCT_NumVal.Clear;
begin
  FAssigneds := [];
  FV := '';
  FIdx := 2147483632;
  FFormatCode := '';
end;

procedure TCT_NumVal.Assign(AItem: TCT_NumVal);
begin
end;

procedure TCT_NumVal.CopyTo(AItem: TCT_NumVal);
begin
end;

{ TCT_NumValXpgList }

function  TCT_NumValXpgList.GetItems(Index: integer): TCT_NumVal;
begin
  Result := TCT_NumVal(inherited Items[Index]);
end;

function  TCT_NumValXpgList.Add: TCT_NumVal;
begin
  Result := TCT_NumVal.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_NumValXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_NumValXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_NumValXpgList.Assign(AItem: TCT_NumValXpgList);
begin
end;

procedure TCT_NumValXpgList.CopyTo(AItem: TCT_NumValXpgList);
begin
end;

{ TCT_Layout }

function  TCT_Layout.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FManualLayout <> Nil then 
    Inc(ElemsAssigned,FManualLayout.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_Layout.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $00000599: begin
      if FManualLayout = Nil then
        FManualLayout := TCT_ManualLayout.Create(FOwner);
      Result := FManualLayout;
    end;
    $00000321: begin
      if FExtLst = Nil then
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then
    Result.Assigneds := [xaRead];
end;

procedure TCT_Layout.Prepare(const AX,AY,AW,AH: double);
begin
  if FManualLayout <> NIl then begin
    FX := FManualLayout.X.Val;
    FY := FManualLayout.Y.Val;
    FW := FManualLayout.W.Val;
    FH := FManualLayout.H.Val;
  end
  else begin
    FX := AX;
    FY := AY;
    FW := AW;
    FH := AH;
  end;
end;

procedure TCT_Layout.Write(AWriter: TXpgWriteXML);
begin
  if (FManualLayout <> Nil) and FManualLayout.Assigned then 
    if xaElements in FManualLayout.FAssigneds then 
    begin
      AWriter.BeginTag('c:manualLayout');
      FManualLayout.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:manualLayout');
{$ifdef XLS_WRITE_EXTLST}
  if (FExtLst <> Nil) and FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then
    begin
      AWriter.BeginTag('c:extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:extLst');
{$endif}
end;

constructor TCT_Layout.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 0;
end;

destructor TCT_Layout.Destroy;
begin
  if FManualLayout <> Nil then 
    FManualLayout.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_Layout.Clear;
begin
  FAssigneds := [];
  if FManualLayout <> Nil then 
    FreeAndNil(FManualLayout);
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
end;

procedure TCT_Layout.Assign(AItem: TCT_Layout);
begin
end;

procedure TCT_Layout.CopyTo(AItem: TCT_Layout);
begin
end;

function  TCT_Layout.Create_ManualLayout: TCT_ManualLayout;
begin
  Result := TCT_ManualLayout.Create(FOwner);
  FManualLayout := Result;
end;

function  TCT_Layout.Create_ExtLst: TCT_ExtensionList;
begin
  Result := TCT_ExtensionList.Create(FOwner);
  FExtLst := Result;
end;

{ TCT_Tx }

function  TCT_Tx.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FStrRef <> Nil then 
    Inc(ElemsAssigned,FStrRef.CheckAssigned);
  if FRich <> Nil then 
    Inc(ElemsAssigned,FRich.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_Tx.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $00000313: begin
      if FStrRef = Nil then 
        FStrRef := TCT_StrRef.Create(FOwner);
      Result := FStrRef;
    end;
    $00000243: begin
      if FRich = Nil then 
        FRich := TCT_TextBody.Create(FOwner);
      Result := FRich;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_Tx.Write(AWriter: TXpgWriteXML);
begin
  if (FStrRef <> Nil) and FStrRef.Assigned then 
    if xaElements in FStrRef.FAssigneds then 
    begin
      AWriter.BeginTag('c:strRef');
      FStrRef.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:strRef');
// TODO
//  else
//    AWriter.SimpleTag('c:strRef');
  if (FRich <> Nil) and FRich.Assigned then
    if xaElements in FRich.Assigneds then
    begin
      AWriter.BeginTag('c:rich');
      FRich.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:rich')
  else 
    AWriter.SimpleTag('c:rich');
end;

constructor TCT_Tx.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 0;
end;

destructor TCT_Tx.Destroy;
begin
  if FStrRef <> Nil then 
    FStrRef.Free;
  if FRich <> Nil then 
    FRich.Free;
end;

procedure TCT_Tx.Clear;
begin
  FAssigneds := [];
  if FStrRef <> Nil then 
    FreeAndNil(FStrRef);
  if FRich <> Nil then 
    FreeAndNil(FRich);
end;

procedure TCT_Tx.Assign(AItem: TCT_Tx);
begin
end;

procedure TCT_Tx.CopyTo(AItem: TCT_Tx);
begin
end;

function  TCT_Tx.Create_StrRef: TCT_StrRef;
begin
  Result := TCT_StrRef.Create(FOwner);
  FStrRef := Result;
end;

function  TCT_Tx.Create_Rich: TCT_TextBody;
begin
  Result := TCT_TextBody.Create(FOwner);
  FRich := Result;
end;

{ TEG_DLblShared }

function  TEG_DLblShared.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FNumFmt <> Nil then 
    Inc(ElemsAssigned,FNumFmt.CheckAssigned);
  if FSpPr <> Nil then 
    Inc(ElemsAssigned,FSpPr.CheckAssigned);
  if FTxPr <> Nil then 
    Inc(ElemsAssigned,FTxPr.CheckAssigned);
  if FDLblPos <> Nil then 
    Inc(ElemsAssigned,FDLblPos.CheckAssigned);
  if FShowLegendKey <> Nil then 
    Inc(ElemsAssigned,FShowLegendKey.CheckAssigned);
  if FShowVal <> Nil then 
    Inc(ElemsAssigned,FShowVal.CheckAssigned);
  if FShowCatName <> Nil then 
    Inc(ElemsAssigned,FShowCatName.CheckAssigned);
  if FShowSerName <> Nil then 
    Inc(ElemsAssigned,FShowSerName.CheckAssigned);
  if FShowPercent <> Nil then 
    Inc(ElemsAssigned,FShowPercent.CheckAssigned);
  if FShowBubbleSize <> Nil then 
    Inc(ElemsAssigned,FShowBubbleSize.CheckAssigned);
  if FSeparator <> '' then 
    Inc(ElemsAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TEG_DLblShared.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $00000314: begin
      if FNumFmt = Nil then 
        FNumFmt := TCT_NumFmt.Create(FOwner);
      Result := FNumFmt;
    end;
    $00000242: begin
      if FSpPr = Nil then 
        FSpPr := TCT_ShapeProperties.Create(FOwner);
      Result := FSpPr;
    end;
    $0000024B: begin
      if FTxPr = Nil then 
        FTxPr := TCT_TextBody.Create(FOwner);
      Result := FTxPr;
    end;
    $0000034D: begin
      if FDLblPos = Nil then 
        FDLblPos := TCT_DLblPos.Create(FOwner);
      Result := FDLblPos;
    end;
    $000005D6: begin
      if FShowLegendKey = Nil then 
        FShowLegendKey := TCT_Boolean.Create(FOwner);
      Result := FShowLegendKey;
    end;
    $00000381: begin
      if FShowVal = Nil then 
        FShowVal := TCT_Boolean.Create(FOwner);
      Result := FShowVal;
    end;
    $000004F7: begin
      if FShowCatName = Nil then 
        FShowCatName := TCT_Boolean.Create(FOwner);
      Result := FShowCatName;
    end;
    $00000509: begin
      if FShowSerName = Nil then 
        FShowSerName := TCT_Boolean.Create(FOwner);
      Result := FShowSerName;
    end;
    $0000052F: begin
      if FShowPercent = Nil then 
        FShowPercent := TCT_Boolean.Create(FOwner);
      Result := FShowPercent;
    end;
    $00000645: begin
      if FShowBubbleSize = Nil then 
        FShowBubbleSize := TCT_Boolean.Create(FOwner);
      Result := FShowBubbleSize;
    end;
    $0000046E: FSeparator := AReader.Text;
    else 
    begin
      Result := Nil;
      Exit;
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TEG_DLblShared.Write(AWriter: TXpgWriteXML);
begin
  if (FNumFmt <> Nil) and FNumFmt.Assigned then 
  begin
    FNumFmt.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:numFmt');
  end;
  if (FSpPr <> Nil) and FSpPr.Assigned then 
  begin
    FSpPr.WriteAttributes(AWriter);
    if xaElements in FSpPr.Assigneds then
    begin
      AWriter.BeginTag('c:spPr');
      FSpPr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:spPr');
  end;
  if (FTxPr <> Nil) and FTxPr.Assigned then 
    if xaElements in FTxPr.Assigneds then
    begin
      AWriter.BeginTag('c:txPr');
      FTxPr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:txPr');
  if (FDLblPos <> Nil) and FDLblPos.Assigned then 
  begin
    FDLblPos.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:dLblPos');
  end;
  if (FShowLegendKey <> Nil) and FShowLegendKey.Assigned then 
  begin
    FShowLegendKey.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:showLegendKey');
  end;
  if (FShowVal <> Nil) and FShowVal.Assigned then 
  begin
    FShowVal.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:showVal');
  end;
  if (FShowCatName <> Nil) and FShowCatName.Assigned then 
  begin
    FShowCatName.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:showCatName');
  end;
  if (FShowSerName <> Nil) and FShowSerName.Assigned then 
  begin
    FShowSerName.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:showSerName');
  end;
  if (FShowPercent <> Nil) and FShowPercent.Assigned then 
  begin
    FShowPercent.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:showPercent');
  end;
  if (FShowBubbleSize <> Nil) and FShowBubbleSize.Assigned then 
  begin
    FShowBubbleSize.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:showBubbleSize');
  end;
  if FSeparator <> '' then 
    AWriter.SimpleTextTag('c:separator',FSeparator);
end;

constructor TEG_DLblShared.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 11;
  FAttributeCount := 0;
end;

destructor TEG_DLblShared.Destroy;
begin
  if FNumFmt <> Nil then 
    FNumFmt.Free;
  if FSpPr <> Nil then 
    FSpPr.Free;
  if FTxPr <> Nil then 
    FTxPr.Free;
  if FDLblPos <> Nil then 
    FDLblPos.Free;
  if FShowLegendKey <> Nil then 
    FShowLegendKey.Free;
  if FShowVal <> Nil then 
    FShowVal.Free;
  if FShowCatName <> Nil then 
    FShowCatName.Free;
  if FShowSerName <> Nil then 
    FShowSerName.Free;
  if FShowPercent <> Nil then 
    FShowPercent.Free;
  if FShowBubbleSize <> Nil then 
    FShowBubbleSize.Free;
end;

procedure TEG_DLblShared.Clear;
begin
  FAssigneds := [];
  if FNumFmt <> Nil then 
    FreeAndNil(FNumFmt);
  if FSpPr <> Nil then 
    FreeAndNil(FSpPr);
  if FTxPr <> Nil then 
    FreeAndNil(FTxPr);
  if FDLblPos <> Nil then 
    FreeAndNil(FDLblPos);
  if FShowLegendKey <> Nil then 
    FreeAndNil(FShowLegendKey);
  if FShowVal <> Nil then 
    FreeAndNil(FShowVal);
  if FShowCatName <> Nil then 
    FreeAndNil(FShowCatName);
  if FShowSerName <> Nil then 
    FreeAndNil(FShowSerName);
  if FShowPercent <> Nil then 
    FreeAndNil(FShowPercent);
  if FShowBubbleSize <> Nil then 
    FreeAndNil(FShowBubbleSize);
  FSeparator := '';
end;

procedure TEG_DLblShared.Assign(AItem: TEG_DLblShared);
begin
end;

procedure TEG_DLblShared.CopyTo(AItem: TEG_DLblShared);
begin
end;

function  TEG_DLblShared.Create_NumFmt: TCT_NumFmt;
begin
  Result := TCT_NumFmt.Create(FOwner);
  FNumFmt := Result;
end;

function  TEG_DLblShared.Create_SpPr: TCT_ShapeProperties;
begin
  Result := TCT_ShapeProperties.Create(FOwner);
  FSpPr := Result;
end;

function  TEG_DLblShared.Create_TxPr: TCT_TextBody;
begin
  Result := TCT_TextBody.Create(FOwner);
  FTxPr := Result;
end;

function  TEG_DLblShared.Create_DLblPos: TCT_DLblPos;
begin
  Result := TCT_DLblPos.Create(FOwner);
  FDLblPos := Result;
end;

function  TEG_DLblShared.Create_ShowLegendKey: TCT_Boolean;
begin
  Result := TCT_Boolean.Create(FOwner);
  FShowLegendKey := Result;
end;

function  TEG_DLblShared.Create_ShowVal: TCT_Boolean;
begin
  Result := TCT_Boolean.Create(FOwner);
  FShowVal := Result;
end;

function  TEG_DLblShared.Create_ShowCatName: TCT_Boolean;
begin
  Result := TCT_Boolean.Create(FOwner);
  FShowCatName := Result;
end;

function  TEG_DLblShared.Create_ShowSerName: TCT_Boolean;
begin
  Result := TCT_Boolean.Create(FOwner);
  FShowSerName := Result;
end;

function  TEG_DLblShared.Create_ShowPercent: TCT_Boolean;
begin
  Result := TCT_Boolean.Create(FOwner);
  FShowPercent := Result;
end;

function  TEG_DLblShared.Create_ShowBubbleSize: TCT_Boolean;
begin
  Result := TCT_Boolean.Create(FOwner);
  FShowBubbleSize := Result;
end;

{ TCT_NumData }

function  TCT_NumData.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FFormatCode <> '' then 
    Inc(ElemsAssigned);
  if FPtCount <> Nil then 
    Inc(ElemsAssigned,FPtCount.CheckAssigned);
  if FPtXpgList <> Nil then 
    Inc(ElemsAssigned,FPtXpgList.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_NumData.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $000004A1: FFormatCode := AReader.Text;
    $0000038A: begin
      if FPtCount = Nil then 
        FPtCount := TCT_UnsignedInt.Create(FOwner);
      Result := FPtCount;
    end;
    $00000181: begin
      if FPtXpgList = Nil then 
        FPtXpgList := TCT_NumValXpgList.Create(FOwner);
      Result := FPtXpgList.Add;
    end;
    $00000321: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_NumData.Write(AWriter: TXpgWriteXML);
begin
  if FFormatCode <> '' then 
    AWriter.SimpleTextTag('c:formatCode',FFormatCode);
  if (FPtCount <> Nil) and FPtCount.Assigned then 
  begin
    FPtCount.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:ptCount');
  end;
  if FPtXpgList <> Nil then 
    FPtXpgList.Write(AWriter,'c:pt');
{$ifdef XLS_WRITE_EXTLST}
  if (FExtLst <> Nil) and FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then
    begin
      AWriter.BeginTag('c:extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:extLst');
{$endif}
end;

constructor TCT_NumData.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 4;
  FAttributeCount := 0;
end;

destructor TCT_NumData.Destroy;
begin
  if FPtCount <> Nil then 
    FPtCount.Free;
  if FPtXpgList <> Nil then 
    FPtXpgList.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_NumData.Clear;
begin
  FAssigneds := [];
  FFormatCode := '';
  if FPtCount <> Nil then 
    FreeAndNil(FPtCount);
  if FPtXpgList <> Nil then 
    FreeAndNil(FPtXpgList);
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
end;

procedure TCT_NumData.Assign(AItem: TCT_NumData);
begin
end;

procedure TCT_NumData.CopyTo(AItem: TCT_NumData);
begin
end;

function  TCT_NumData.Create_PtCount: TCT_UnsignedInt;
begin
  Result := TCT_UnsignedInt.Create(FOwner);
  FPtCount := Result;
end;

function  TCT_NumData.Create_PtXpgList: TCT_NumValXpgList;
begin
  Result := TCT_NumValXpgList.Create(FOwner);
  FPtXpgList := Result;
end;

function  TCT_NumData.Create_ExtLst: TCT_ExtensionList;
begin
  Result := TCT_ExtensionList.Create(FOwner);
  FExtLst := Result;
end;

{ TCT_Lvl }

function  TCT_Lvl.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FPtXpgList <> Nil then 
    Inc(ElemsAssigned,FPtXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_Lvl.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'c:pt' then 
  begin
    if FPtXpgList = Nil then 
      FPtXpgList := TCT_StrValXpgList.Create(FOwner);
    Result := FPtXpgList.Add;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_Lvl.Write(AWriter: TXpgWriteXML);
begin
  if FPtXpgList <> Nil then 
    FPtXpgList.Write(AWriter,'c:pt');
end;

constructor TCT_Lvl.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
end;

destructor TCT_Lvl.Destroy;
begin
  if FPtXpgList <> Nil then 
    FPtXpgList.Free;
end;

procedure TCT_Lvl.Clear;
begin
  FAssigneds := [];
  if FPtXpgList <> Nil then 
    FreeAndNil(FPtXpgList);
end;

procedure TCT_Lvl.Assign(AItem: TCT_Lvl);
begin
end;

procedure TCT_Lvl.CopyTo(AItem: TCT_Lvl);
begin
end;

function  TCT_Lvl.Create_PtXpgList: TCT_StrValXpgList;
begin
  Result := TCT_StrValXpgList.Create(FOwner);
  FPtXpgList := Result;
end;

{ TCT_LvlXpgList }

function  TCT_LvlXpgList.GetItems(Index: integer): TCT_Lvl;
begin
  Result := TCT_Lvl(inherited Items[Index]);
end;

function  TCT_LvlXpgList.Add: TCT_Lvl;
begin
  Result := TCT_Lvl.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_LvlXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_LvlXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
    if xaElements in Items[i].Assigneds then
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
end;

procedure TCT_LvlXpgList.Assign(AItem: TCT_LvlXpgList);
begin
end;

procedure TCT_LvlXpgList.CopyTo(AItem: TCT_LvlXpgList);
begin
end;

{ TCT_MarkerStyle }

function  TCT_MarkerStyle.CheckAssigned: integer;
begin
  FAssigneds := [xaAttributes];
  if FVal <> TST_MarkerStyle(XPG_UNKNOWN_ENUM) then
    Result := 1
  else
    Result := 0;
end;

procedure TCT_MarkerStyle.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_MarkerStyle.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FVal <> TST_MarkerStyle(XPG_UNKNOWN_ENUM) then
    AWriter.AddAttribute('val',StrTST_MarkerStyle[Integer(FVal)]);
end;

procedure TCT_MarkerStyle.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then
    FVal := TST_MarkerStyle(StrToEnum('stms' + AAttributes.Values[0]))
  else
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_MarkerStyle.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  Clear;
end;

destructor TCT_MarkerStyle.Destroy;
begin
end;

procedure TCT_MarkerStyle.Clear;
begin
  FAssigneds := [];
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := TST_MarkerStyle(XPG_UNKNOWN_ENUM);
end;

procedure TCT_MarkerStyle.Assign(AItem: TCT_MarkerStyle);
begin
end;

procedure TCT_MarkerStyle.CopyTo(AItem: TCT_MarkerStyle);
begin
end;

{ TCT_MarkerSize }

function  TCT_MarkerSize.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FVal <> MAXINT then
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then
    FAssigneds := [xaAttributes];
end;

procedure TCT_MarkerSize.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_MarkerSize.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FVal <> MAXINT then
    AWriter.AddAttribute('val',XmlIntToStr(FVal));
end;

procedure TCT_MarkerSize.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then
    FVal := XmlStrToIntDef(AAttributes.Values[0],0)
  else
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_MarkerSize.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := MAXINT;
end;

destructor TCT_MarkerSize.Destroy;
begin
end;

procedure TCT_MarkerSize.Clear;
begin
  FAssigneds := [];
  FVal := MAXINT;
end;

procedure TCT_MarkerSize.Assign(AItem: TCT_MarkerSize);
begin
end;

procedure TCT_MarkerSize.CopyTo(AItem: TCT_MarkerSize);
begin
end;

{ TCT_PictureFormat }

function  TCT_PictureFormat.CheckAssigned: integer;
begin
  FAssigneds := [xaAttributes];
  if FVal <> TST_PictureFormat(XPG_UNKNOWN_ENUM) then
    Result := 1
  else
    Result := 0;
end;

procedure TCT_PictureFormat.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_PictureFormat.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FVal <> TST_PictureFormat(XPG_UNKNOWN_ENUM) then
    AWriter.AddAttribute('val',StrTST_PictureFormat[Integer(FVal)]);
end;

procedure TCT_PictureFormat.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then
    FVal := TST_PictureFormat(StrToEnum('stpf' + AAttributes.Values[0]))
  else
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_PictureFormat.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  Clear;
end;

destructor TCT_PictureFormat.Destroy;
begin
end;

procedure TCT_PictureFormat.Clear;
begin
  FAssigneds := [];
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := TST_PictureFormat(XPG_UNKNOWN_ENUM);
end;

procedure TCT_PictureFormat.Assign(AItem: TCT_PictureFormat);
begin
end;

procedure TCT_PictureFormat.CopyTo(AItem: TCT_PictureFormat);
begin
end;

{ TCT_PictureStackUnit }

function  TCT_PictureStackUnit.CheckAssigned: integer;
begin
  FAssigneds := [xaAttributes];
  Result := 1;
end;

procedure TCT_PictureStackUnit.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_PictureStackUnit.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('val',XmlFloatToStr(FVal));
end;

procedure TCT_PictureStackUnit.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then 
    FVal := XmlStrToFloatDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_PictureStackUnit.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := NaN;
end;

destructor TCT_PictureStackUnit.Destroy;
begin
end;

procedure TCT_PictureStackUnit.Clear;
begin
  FAssigneds := [];
  FVal := NaN;
end;

procedure TCT_PictureStackUnit.Assign(AItem: TCT_PictureStackUnit);
begin
end;

procedure TCT_PictureStackUnit.CopyTo(AItem: TCT_PictureStackUnit);
begin
end;

{ TGroup_DLbl }

function  TGroup_DLbl.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FLayout <> Nil then 
    Inc(ElemsAssigned,FLayout.CheckAssigned);
  if FTx <> Nil then 
    Inc(ElemsAssigned,FTx.CheckAssigned);
  Inc(ElemsAssigned,FEG_DLblShared.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TGroup_DLbl.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $0000033B: begin
      if FLayout = Nil then 
        FLayout := TCT_Layout.Create(FOwner);
      Result := FLayout;
    end;
    $00000189: begin
      if FTx = Nil then 
        FTx := TCT_Tx.Create(FOwner);
      Result := FTx;
    end;
    else 
    begin
      Result := FEG_DLblShared.HandleElement(AReader);
      if Result = Nil then 
        FOwner.Errors.Error(xemUnknownElement,AReader.QName);
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TGroup_DLbl.Write(AWriter: TXpgWriteXML);
begin
  if (FLayout <> Nil) and FLayout.Assigned then 
    if xaElements in FLayout.Assigneds then
    begin
      AWriter.BeginTag('c:layout');
      FLayout.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:layout');
  if (FTx <> Nil) and FTx.Assigned then 
    if xaElements in FTx.Assigneds then
    begin
      AWriter.BeginTag('c:tx');
      FTx.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:tx');
  FEG_DLblShared.Write(AWriter);
end;

constructor TGroup_DLbl.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 3;
  FAttributeCount := 0;
  FEG_DLblShared := TEG_DLblShared.Create(FOwner);
end;

destructor TGroup_DLbl.Destroy;
begin
  if FLayout <> Nil then 
    FLayout.Free;
  if FTx <> Nil then 
    FTx.Free;
  FEG_DLblShared.Free;
end;

procedure TGroup_DLbl.Clear;
begin
  FAssigneds := [];
  if FLayout <> Nil then 
    FreeAndNil(FLayout);
  if FTx <> Nil then 
    FreeAndNil(FTx);
  FEG_DLblShared.Clear;
end;

procedure TGroup_DLbl.Assign(AItem: TGroup_DLbl);
begin
end;

procedure TGroup_DLbl.CopyTo(AItem: TGroup_DLbl);
begin
end;

function  TGroup_DLbl.Create_Layout: TCT_Layout;
begin
  Result := TCT_Layout.Create(FOwner);
  FLayout := Result;
end;

function  TGroup_DLbl.Create_Tx: TCT_Tx;
begin
  Result := TCT_Tx.Create(FOwner);
  FTx := Result;
end;

{ TCT_ChartLines }

function  TCT_ChartLines.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FSpPr <> Nil then 
    Inc(ElemsAssigned,FSpPr.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_ChartLines.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'c:spPr' then 
  begin
    if FSpPr = Nil then 
      FSpPr := TCT_ShapeProperties.Create(FOwner);
    Result := FSpPr;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_ChartLines.Write(AWriter: TXpgWriteXML);
begin
  if (FSpPr <> Nil) and FSpPr.Assigned then 
  begin
    FSpPr.WriteAttributes(AWriter);
    if xaElements in FSpPr.Assigneds then
    begin
      AWriter.BeginTag('c:spPr');
      FSpPr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:spPr');
  end;
end;

constructor TCT_ChartLines.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
end;

destructor TCT_ChartLines.Destroy;
begin
  if FSpPr <> Nil then 
    FSpPr.Free;
end;

procedure TCT_ChartLines.Clear;
begin
  FAssigneds := [];
  if FSpPr <> Nil then 
    FreeAndNil(FSpPr);
end;

procedure TCT_ChartLines.Assign(AItem: TCT_ChartLines);
begin
end;

procedure TCT_ChartLines.CopyTo(AItem: TCT_ChartLines);
begin
end;

function  TCT_ChartLines.Create_SpPr: TCT_ShapeProperties;
begin
  Result := TCT_ShapeProperties.Create(FOwner);
  FSpPr := Result;
end;

{ TCT_ChartLinesXpgList }

function  TCT_ChartLinesXpgList.GetItems(Index: integer): TCT_ChartLines;
begin
  Result := TCT_ChartLines(inherited Items[Index]);
end;

function  TCT_ChartLinesXpgList.Add: TCT_ChartLines;
begin
  Result := TCT_ChartLines.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_ChartLinesXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_ChartLinesXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
    if xaElements in Items[i].Assigneds then
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
end;

procedure TCT_ChartLinesXpgList.Assign(AItem: TCT_ChartLinesXpgList);
begin
end;

procedure TCT_ChartLinesXpgList.CopyTo(AItem: TCT_ChartLinesXpgList);
begin
end;

{ TCT_NumRef }

function  TCT_NumRef.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FF <> '' then 
    Inc(ElemsAssigned);
  if FNumCache <> Nil then 
    Inc(ElemsAssigned,FNumCache.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_NumRef.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $00000103: FF := AReader.Text;
    $000003C1: begin
      if FNumCache = Nil then 
        FNumCache := TCT_NumData.Create(FOwner);
      Result := FNumCache;
    end;
    $00000321: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_NumRef.Write(AWriter: TXpgWriteXML);
begin
  if FF <> '' then 
    AWriter.SimpleTextTag('c:f',FF);
  if (FNumCache <> Nil) and FNumCache.Assigned then 
    if xaElements in FNumCache.Assigneds then
    begin
      AWriter.BeginTag('c:numCache');
      FNumCache.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:numCache');
{$ifdef XLS_WRITE_EXTLST}
  if (FExtLst <> Nil) and FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then
    begin
      AWriter.BeginTag('c:extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:extLst');
{$endif}
end;

constructor TCT_NumRef.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 3;
  FAttributeCount := 0;
end;

destructor TCT_NumRef.Destroy;
begin
  if FNumCache <> Nil then 
    FNumCache.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_NumRef.Clear;
begin
  FAssigneds := [];
  FF := '';
  if FNumCache <> Nil then 
    FreeAndNil(FNumCache);
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
end;

procedure TCT_NumRef.Assign(AItem: TCT_NumRef);
begin
end;

procedure TCT_NumRef.CopyTo(AItem: TCT_NumRef);
begin
end;

function  TCT_NumRef.Create_NumCache: TCT_NumData;
begin
  Result := TCT_NumData.Create(FOwner);
  FNumCache := Result;
end;

function  TCT_NumRef.Create_ExtLst: TCT_ExtensionList;
begin
  Result := TCT_ExtensionList.Create(FOwner);
  FExtLst := Result;
end;

{ TCT_MultiLvlStrData }

function  TCT_MultiLvlStrData.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FPtCount <> Nil then 
    Inc(ElemsAssigned,FPtCount.CheckAssigned);
  if FLvlXpgList <> Nil then 
    Inc(ElemsAssigned,FLvlXpgList.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_MultiLvlStrData.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $0000038A: begin
      if FPtCount = Nil then 
        FPtCount := TCT_UnsignedInt.Create(FOwner);
      Result := FPtCount;
    end;
    $000001EB: begin
      if FLvlXpgList = Nil then 
        FLvlXpgList := TCT_LvlXpgList.Create(FOwner);
      Result := FLvlXpgList.Add;
    end;
    $00000321: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_MultiLvlStrData.Write(AWriter: TXpgWriteXML);
begin
  if (FPtCount <> Nil) and FPtCount.Assigned then 
  begin
    FPtCount.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:ptCount');
  end;
  if FLvlXpgList <> Nil then 
    FLvlXpgList.Write(AWriter,'c:lvl');
{$ifdef XLS_WRITE_EXTLST}
  if (FExtLst <> Nil) and FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then
    begin
      AWriter.BeginTag('c:extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:extLst');
{$endif}
end;

constructor TCT_MultiLvlStrData.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 3;
  FAttributeCount := 0;
end;

destructor TCT_MultiLvlStrData.Destroy;
begin
  if FPtCount <> Nil then 
    FPtCount.Free;
  if FLvlXpgList <> Nil then 
    FLvlXpgList.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_MultiLvlStrData.Clear;
begin
  FAssigneds := [];
  if FPtCount <> Nil then 
    FreeAndNil(FPtCount);
  if FLvlXpgList <> Nil then 
    FreeAndNil(FLvlXpgList);
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
end;

procedure TCT_MultiLvlStrData.Assign(AItem: TCT_MultiLvlStrData);
begin
end;

procedure TCT_MultiLvlStrData.CopyTo(AItem: TCT_MultiLvlStrData);
begin
end;

function  TCT_MultiLvlStrData.Create_PtCount: TCT_UnsignedInt;
begin
  Result := TCT_UnsignedInt.Create(FOwner);
  FPtCount := Result;
end;

function  TCT_MultiLvlStrData.Create_LvlXpgList: TCT_LvlXpgList;
begin
  Result := TCT_LvlXpgList.Create(FOwner);
  FLvlXpgList := Result;
end;

function  TCT_MultiLvlStrData.Create_ExtLst: TCT_ExtensionList;
begin
  Result := TCT_ExtensionList.Create(FOwner);
  FExtLst := Result;
end;

{ TCT_SerTx }

function  TCT_SerTx.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FStrRef <> Nil then 
    Inc(ElemsAssigned,FStrRef.CheckAssigned);
  if FV <> '' then 
    Inc(ElemsAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_SerTx.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $00000313: begin
      if FStrRef = Nil then 
        FStrRef := TCT_StrRef.Create(FOwner);
      Result := FStrRef;
    end;
    $00000113: FV := AReader.Text;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_SerTx.Write(AWriter: TXpgWriteXML);
begin
  if (FStrRef <> Nil) and FStrRef.Assigned then begin
    if xaElements in FStrRef.Assigneds then
    begin
      AWriter.BeginTag('c:strRef');
      FStrRef.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:strRef');
  end
  else
    AWriter.SimpleTextTag('c:v',FV);
end;

constructor TCT_SerTx.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 0;
end;

destructor TCT_SerTx.Destroy;
begin
  if FStrRef <> Nil then 
    FStrRef.Free;
end;

procedure TCT_SerTx.Clear;
begin
  FAssigneds := [];
  if FStrRef <> Nil then 
    FreeAndNil(FStrRef);
  FV := '';
end;

procedure TCT_SerTx.Assign(AItem: TCT_SerTx);
begin
end;

procedure TCT_SerTx.CopyTo(AItem: TCT_SerTx);
begin
end;

function  TCT_SerTx.Create_StrRef: TCT_StrRef;
begin
  Result := TCT_StrRef.Create(FOwner);
  FStrRef := Result;
end;

{ TCT_PictureOptions }

function  TCT_PictureOptions.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FApplyToFront <> Nil then 
    Inc(ElemsAssigned,FApplyToFront.CheckAssigned);
  if FApplyToSides <> Nil then 
    Inc(ElemsAssigned,FApplyToSides.CheckAssigned);
  if FApplyToEnd <> Nil then 
    Inc(ElemsAssigned,FApplyToEnd.CheckAssigned);
  if FPictureFormat <> Nil then 
    Inc(ElemsAssigned,FPictureFormat.CheckAssigned);
  if FPictureStackUnit <> Nil then 
    Inc(ElemsAssigned,FPictureStackUnit.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_PictureOptions.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $0000058F: begin
      if FApplyToFront = Nil then 
        FApplyToFront := TCT_Boolean.Create(FOwner);
      Result := FApplyToFront;
    end;
    $0000057E: begin
      if FApplyToSides = Nil then 
        FApplyToSides := TCT_Boolean.Create(FOwner);
      Result := FApplyToSides;
    end;
    $0000049D: begin
      if FApplyToEnd = Nil then 
        FApplyToEnd := TCT_Boolean.Create(FOwner);
      Result := FApplyToEnd;
    end;
    $00000602: begin
      if FPictureFormat = Nil then 
        FPictureFormat := TCT_PictureFormat.Create(FOwner);
      Result := FPictureFormat;
    end;
    $0000072F: begin
      if FPictureStackUnit = Nil then 
        FPictureStackUnit := TCT_PictureStackUnit.Create(FOwner);
      Result := FPictureStackUnit;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_PictureOptions.Write(AWriter: TXpgWriteXML);
begin
  if (FApplyToFront <> Nil) and FApplyToFront.Assigned then 
  begin
    FApplyToFront.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:applyToFront');
  end;
  if (FApplyToSides <> Nil) and FApplyToSides.Assigned then 
  begin
    FApplyToSides.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:applyToSides');
  end;
  if (FApplyToEnd <> Nil) and FApplyToEnd.Assigned then 
  begin
    FApplyToEnd.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:applyToEnd');
  end;
  if (FPictureFormat <> Nil) and FPictureFormat.Assigned then 
  begin
    FPictureFormat.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:pictureFormat');
  end;
  if (FPictureStackUnit <> Nil) and FPictureStackUnit.Assigned then 
  begin
    FPictureStackUnit.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:pictureStackUnit');
  end;
end;

constructor TCT_PictureOptions.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 5;
  FAttributeCount := 0;
end;

destructor TCT_PictureOptions.Destroy;
begin
  if FApplyToFront <> Nil then 
    FApplyToFront.Free;
  if FApplyToSides <> Nil then 
    FApplyToSides.Free;
  if FApplyToEnd <> Nil then 
    FApplyToEnd.Free;
  if FPictureFormat <> Nil then 
    FPictureFormat.Free;
  if FPictureStackUnit <> Nil then 
    FPictureStackUnit.Free;
end;

procedure TCT_PictureOptions.Clear;
begin
  FAssigneds := [];
  if FApplyToFront <> Nil then 
    FreeAndNil(FApplyToFront);
  if FApplyToSides <> Nil then 
    FreeAndNil(FApplyToSides);
  if FApplyToEnd <> Nil then 
    FreeAndNil(FApplyToEnd);
  if FPictureFormat <> Nil then 
    FreeAndNil(FPictureFormat);
  if FPictureStackUnit <> Nil then 
    FreeAndNil(FPictureStackUnit);
end;

procedure TCT_PictureOptions.Assign(AItem: TCT_PictureOptions);
begin
end;

procedure TCT_PictureOptions.CopyTo(AItem: TCT_PictureOptions);
begin
end;

function  TCT_PictureOptions.Create_ApplyToFront: TCT_Boolean;
begin
  Result := TCT_Boolean.Create(FOwner);
  FApplyToFront := Result;
end;

function  TCT_PictureOptions.Create_ApplyToSides: TCT_Boolean;
begin
  Result := TCT_Boolean.Create(FOwner);
  FApplyToSides := Result;
end;

function  TCT_PictureOptions.Create_ApplyToEnd: TCT_Boolean;
begin
  Result := TCT_Boolean.Create(FOwner);
  FApplyToEnd := Result;
end;

function  TCT_PictureOptions.Create_PictureFormat: TCT_PictureFormat;
begin
  Result := TCT_PictureFormat.Create(FOwner);
  FPictureFormat := Result;
end;

function  TCT_PictureOptions.Create_PictureStackUnit: TCT_PictureStackUnit;
begin
  Result := TCT_PictureStackUnit.Create(FOwner);
  FPictureStackUnit := Result;
end;

{ TCT_DLbl }

function  TCT_DLbl.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FIdx <> Nil then 
    Inc(ElemsAssigned,FIdx.CheckAssigned);
  if FDelete <> Nil then 
    Inc(ElemsAssigned,FDelete.CheckAssigned);
  Inc(ElemsAssigned,FGroup_DLbl.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_DLbl.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $000001E2: begin
      if FIdx = Nil then 
        FIdx := TCT_UnsignedInt.Create(FOwner);
      Result := FIdx;
    end;
    $00000310: begin
      if FDelete = Nil then 
        FDelete := TCT_Boolean.Create(FOwner);
      Result := FDelete;
    end;
    $00000321: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
    begin
      Result := FGroup_DLbl.HandleElement(AReader);
      if Result = Nil then 
        FOwner.Errors.Error(xemUnknownElement,AReader.QName);
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_DLbl.Write(AWriter: TXpgWriteXML);
begin
  if (FIdx <> Nil) and FIdx.Assigned then 
  begin
    FIdx.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:idx');
  end
  else 
    AWriter.SimpleTag('c:idx');
  if (FDelete <> Nil) and FDelete.Assigned then 
  begin
    FDelete.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:delete');
  end;
  FGroup_DLbl.Write(AWriter);
{$ifdef XLS_WRITE_EXTLST}
  if (FExtLst <> Nil) and FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then
    begin
      AWriter.BeginTag('c:extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:extLst');
{$endif}
end;

constructor TCT_DLbl.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 4;
  FAttributeCount := 0;
  FGroup_DLbl := TGroup_DLbl.Create(FOwner);
end;

destructor TCT_DLbl.Destroy;
begin
  if FIdx <> Nil then 
    FIdx.Free;
  if FDelete <> Nil then 
    FDelete.Free;
  FGroup_DLbl.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_DLbl.Clear;
begin
  FAssigneds := [];
  if FIdx <> Nil then 
    FreeAndNil(FIdx);
  if FDelete <> Nil then 
    FreeAndNil(FDelete);
  FGroup_DLbl.Clear;
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
end;

procedure TCT_DLbl.Assign(AItem: TCT_DLbl);
begin
end;

procedure TCT_DLbl.CopyTo(AItem: TCT_DLbl);
begin
end;

function  TCT_DLbl.Create_Idx: TCT_UnsignedInt;
begin
  Result := TCT_UnsignedInt.Create(FOwner);
  FIdx := Result;
end;

function  TCT_DLbl.Create_Delete: TCT_Boolean;
begin
  Result := TCT_Boolean.Create(FOwner);
  FDelete := Result;
end;

function  TCT_DLbl.Create_ExtLst: TCT_ExtensionList;
begin
  Result := TCT_ExtensionList.Create(FOwner);
  FExtLst := Result;
end;

{ TCT_DLblXpgList }

function  TCT_DLblXpgList.GetItems(Index: integer): TCT_DLbl;
begin
  Result := TCT_DLbl(inherited Items[Index]);
end;

function  TCT_DLblXpgList.Add: TCT_DLbl;
begin
  Result := TCT_DLbl.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_DLblXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_DLblXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
    if xaElements in Items[i].Assigneds then
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
end;

procedure TCT_DLblXpgList.Assign(AItem: TCT_DLblXpgList);
begin
end;

procedure TCT_DLblXpgList.CopyTo(AItem: TCT_DLblXpgList);
begin
end;

{ TGroup_DLbls }

function  TGroup_DLbls.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FEG_DLblShared.CheckAssigned);
  if FShowLeaderLines <> Nil then 
    Inc(ElemsAssigned,FShowLeaderLines.CheckAssigned);
  if FLeaderLines <> Nil then 
    Inc(ElemsAssigned,FLeaderLines.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TGroup_DLbls.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $000006A6: begin
      if FShowLeaderLines = Nil then 
        FShowLeaderLines := TCT_Boolean.Create(FOwner);
      Result := FShowLeaderLines;
    end;
    $00000505: begin
      if FLeaderLines = Nil then 
        FLeaderLines := TCT_ChartLines.Create(FOwner);
      Result := FLeaderLines;
    end;
    else 
    begin
      Result := FEG_DLblShared.HandleElement(AReader);
      if Result = Nil then 
        FOwner.Errors.Error(xemUnknownElement,AReader.QName);
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TGroup_DLbls.Write(AWriter: TXpgWriteXML);
begin
  FEG_DLblShared.Write(AWriter);
  if (FShowLeaderLines <> Nil) and FShowLeaderLines.Assigned then 
  begin
    FShowLeaderLines.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:showLeaderLines');
  end;
  if (FLeaderLines <> Nil) and FLeaderLines.Assigned then 
    if xaElements in FLeaderLines.Assigneds then
    begin
      AWriter.BeginTag('c:leaderLines');
      FLeaderLines.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:leaderLines');
end;

constructor TGroup_DLbls.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 3;
  FAttributeCount := 0;
  FEG_DLblShared := TEG_DLblShared.Create(FOwner);
end;

destructor TGroup_DLbls.Destroy;
begin
  FEG_DLblShared.Free;
  if FShowLeaderLines <> Nil then 
    FShowLeaderLines.Free;
  if FLeaderLines <> Nil then 
    FLeaderLines.Free;
end;

procedure TGroup_DLbls.Clear;
begin
  FAssigneds := [];
  FEG_DLblShared.Clear;
  if FShowLeaderLines <> Nil then 
    FreeAndNil(FShowLeaderLines);
  if FLeaderLines <> Nil then 
    FreeAndNil(FLeaderLines);
end;

procedure TGroup_DLbls.Assign(AItem: TGroup_DLbls);
begin
end;

procedure TGroup_DLbls.CopyTo(AItem: TGroup_DLbls);
begin
end;

function  TGroup_DLbls.Create_ShowLeaderLines: TCT_Boolean;
begin
  Result := TCT_Boolean.Create(FOwner);
  FShowLeaderLines := Result;
end;

function  TGroup_DLbls.Create_LeaderLines: TCT_ChartLines;
begin
  Result := TCT_ChartLines.Create(FOwner);
  FLeaderLines := Result;
end;

{ TCT_TrendlineType }

function  TCT_TrendlineType.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FVal <> stttLinear then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_TrendlineType.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_TrendlineType.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('val',StrTST_TrendlineType[Integer(FVal)]);
end;

procedure TCT_TrendlineType.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then 
    FVal := TST_TrendlineType(StrToEnum('sttt' + AAttributes.Values[0]))
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_TrendlineType.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := stttLinear;
end;

destructor TCT_TrendlineType.Destroy;
begin
end;

procedure TCT_TrendlineType.Clear;
begin
  FAssigneds := [];
  FVal := stttLinear;
end;

procedure TCT_TrendlineType.Assign(AItem: TCT_TrendlineType);
begin
end;

procedure TCT_TrendlineType.CopyTo(AItem: TCT_TrendlineType);
begin
end;

{ TCT_Order }

function  TCT_Order.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FVal <> 2 then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_Order.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_Order.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FVal <> 2 then 
    AWriter.AddAttribute('val',XmlIntToStr(FVal));
end;

procedure TCT_Order.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then 
    FVal := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_Order.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := 2;
end;

destructor TCT_Order.Destroy;
begin
end;

procedure TCT_Order.Clear;
begin
  FAssigneds := [];
  FVal := 2;
end;

procedure TCT_Order.Assign(AItem: TCT_Order);
begin
end;

procedure TCT_Order.CopyTo(AItem: TCT_Order);
begin
end;

{ TCT_Period }

function  TCT_Period.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FVal <> 2 then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_Period.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_Period.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FVal <> 2 then 
    AWriter.AddAttribute('val',XmlIntToStr(FVal));
end;

procedure TCT_Period.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then 
    FVal := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_Period.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := 2;
end;

destructor TCT_Period.Destroy;
begin
end;

procedure TCT_Period.Clear;
begin
  FAssigneds := [];
  FVal := 2;
end;

procedure TCT_Period.Assign(AItem: TCT_Period);
begin
end;

procedure TCT_Period.CopyTo(AItem: TCT_Period);
begin
end;

{ TCT_TrendlineLbl }

function  TCT_TrendlineLbl.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FLayout <> Nil then 
    Inc(ElemsAssigned,FLayout.CheckAssigned);
  if FTx <> Nil then 
    Inc(ElemsAssigned,FTx.CheckAssigned);
  if FNumFmt <> Nil then 
    Inc(ElemsAssigned,FNumFmt.CheckAssigned);
  if FSpPr <> Nil then 
    Inc(ElemsAssigned,FSpPr.CheckAssigned);
  if FTxPr <> Nil then 
    Inc(ElemsAssigned,FTxPr.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_TrendlineLbl.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $0000033B: begin
      if FLayout = Nil then 
        FLayout := TCT_Layout.Create(FOwner);
      Result := FLayout;
    end;
    $00000189: begin
      if FTx = Nil then 
        FTx := TCT_Tx.Create(FOwner);
      Result := FTx;
    end;
    $00000314: begin
      if FNumFmt = Nil then 
        FNumFmt := TCT_NumFmt.Create(FOwner);
      Result := FNumFmt;
    end;
    $00000242: begin
      if FSpPr = Nil then 
        FSpPr := TCT_ShapeProperties.Create(FOwner);
      Result := FSpPr;
    end;
    $0000024B: begin
      if FTxPr = Nil then 
        FTxPr := TCT_TextBody.Create(FOwner);
      Result := FTxPr;
    end;
    $00000321: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_TrendlineLbl.Write(AWriter: TXpgWriteXML);
begin
  if (FLayout <> Nil) and FLayout.Assigned then 
    if xaElements in FLayout.Assigneds then
    begin
      AWriter.BeginTag('c:layout');
      FLayout.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:layout');
  if (FTx <> Nil) and FTx.Assigned then 
    if xaElements in FTx.Assigneds then
    begin
      AWriter.BeginTag('c:tx');
      FTx.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:tx');
  if (FNumFmt <> Nil) and FNumFmt.Assigned then 
  begin
    FNumFmt.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:numFmt');
  end;
  if (FSpPr <> Nil) and FSpPr.Assigned then 
  begin
    FSpPr.WriteAttributes(AWriter);
    if xaElements in FSpPr.Assigneds then
    begin
      AWriter.BeginTag('c:spPr');
      FSpPr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:spPr');
  end;
  if (FTxPr <> Nil) and FTxPr.Assigned then 
    if xaElements in FTxPr.Assigneds then
    begin
      AWriter.BeginTag('c:txPr');
      FTxPr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:txPr');
{$ifdef XLS_WRITE_EXTLST}
  if (FExtLst <> Nil) and FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then
    begin
      AWriter.BeginTag('c:extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:extLst');
{$endif}
end;

constructor TCT_TrendlineLbl.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 6;
  FAttributeCount := 0;
end;

destructor TCT_TrendlineLbl.Destroy;
begin
  if FLayout <> Nil then 
    FLayout.Free;
  if FTx <> Nil then 
    FTx.Free;
  if FNumFmt <> Nil then 
    FNumFmt.Free;
  if FSpPr <> Nil then 
    FSpPr.Free;
  if FTxPr <> Nil then 
    FTxPr.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_TrendlineLbl.Clear;
begin
  FAssigneds := [];
  if FLayout <> Nil then 
    FreeAndNil(FLayout);
  if FTx <> Nil then 
    FreeAndNil(FTx);
  if FNumFmt <> Nil then 
    FreeAndNil(FNumFmt);
  if FSpPr <> Nil then 
    FreeAndNil(FSpPr);
  if FTxPr <> Nil then 
    FreeAndNil(FTxPr);
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
end;

procedure TCT_TrendlineLbl.Assign(AItem: TCT_TrendlineLbl);
begin
end;

procedure TCT_TrendlineLbl.CopyTo(AItem: TCT_TrendlineLbl);
begin
end;

function  TCT_TrendlineLbl.Create_Layout: TCT_Layout;
begin
  Result := TCT_Layout.Create(FOwner);
  FLayout := Result;
end;

function  TCT_TrendlineLbl.Create_Tx: TCT_Tx;
begin
  Result := TCT_Tx.Create(FOwner);
  FTx := Result;
end;

function  TCT_TrendlineLbl.Create_NumFmt: TCT_NumFmt;
begin
  Result := TCT_NumFmt.Create(FOwner);
  FNumFmt := Result;
end;

function  TCT_TrendlineLbl.Create_SpPr: TCT_ShapeProperties;
begin
  Result := TCT_ShapeProperties.Create(FOwner);
  FSpPr := Result;
end;

function  TCT_TrendlineLbl.Create_TxPr: TCT_TextBody;
begin
  Result := TCT_TextBody.Create(FOwner);
  FTxPr := Result;
end;

function  TCT_TrendlineLbl.Create_ExtLst: TCT_ExtensionList;
begin
  Result := TCT_ExtensionList.Create(FOwner);
  FExtLst := Result;
end;

{ TCT_ErrDir }

function  TCT_ErrDir.CheckAssigned: integer;
begin
  FAssigneds := [xaAttributes];
  if FVal <> TST_ErrDir(XPG_UNKNOWN_ENUM) then
    Result := 1
  else
    Result := 0;
end;

procedure TCT_ErrDir.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_ErrDir.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FVal <> TST_ErrDir(XPG_UNKNOWN_ENUM) then
    AWriter.AddAttribute('val',StrTST_ErrDir[Integer(FVal)]);
end;

procedure TCT_ErrDir.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then
    FVal := TST_ErrDir(StrToEnum('sted' + AAttributes.Values[0]))
  else
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_ErrDir.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  Clear;
end;

destructor TCT_ErrDir.Destroy;
begin
end;

procedure TCT_ErrDir.Clear;
begin
  FAssigneds := [];
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := TST_ErrDir(XPG_UNKNOWN_ENUM);
end;

procedure TCT_ErrDir.Assign(AItem: TCT_ErrDir);
begin
end;

procedure TCT_ErrDir.CopyTo(AItem: TCT_ErrDir);
begin
end;

{ TCT_ErrBarType }

function  TCT_ErrBarType.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FVal <> stebtBoth then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_ErrBarType.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_ErrBarType.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FVal <> stebtBoth then 
    AWriter.AddAttribute('val',StrTST_ErrBarType[Integer(FVal)]);
end;

procedure TCT_ErrBarType.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then 
    FVal := TST_ErrBarType(StrToEnum('stebt' + AAttributes.Values[0]))
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_ErrBarType.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := stebtBoth;
end;

destructor TCT_ErrBarType.Destroy;
begin
end;

procedure TCT_ErrBarType.Clear;
begin
  FAssigneds := [];
  FVal := stebtBoth;
end;

procedure TCT_ErrBarType.Assign(AItem: TCT_ErrBarType);
begin
end;

procedure TCT_ErrBarType.CopyTo(AItem: TCT_ErrBarType);
begin
end;

{ TCT_ErrValType }

function  TCT_ErrValType.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FVal <> stevtFixedVal then
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then
    FAssigneds := [xaAttributes];
end;

procedure TCT_ErrValType.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_ErrValType.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FVal <> stevtFixedVal then
    AWriter.AddAttribute('val',StrTST_ErrValType[Integer(FVal)]);
end;

procedure TCT_ErrValType.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then
    FVal := TST_ErrValType(StrToEnum('stevt' + AAttributes.Values[0]))
  else
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_ErrValType.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := stevtFixedVal;
end;

destructor TCT_ErrValType.Destroy;
begin
end;

procedure TCT_ErrValType.Clear;
begin
  FAssigneds := [];
  FVal := stevtFixedVal;
end;

{ TCT_NumDataSource }

function  TCT_NumDataSource.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FNumRef <> Nil then 
    Inc(ElemsAssigned,FNumRef.CheckAssigned);
  if FNumLit <> Nil then 
    Inc(ElemsAssigned,FNumLit.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_NumDataSource.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $0000030A: begin
      if FNumRef = Nil then 
        FNumRef := TCT_NumRef.Create(FOwner);
      Result := FNumRef;
    end;
    $00000316: begin
      if FNumLit = Nil then 
        FNumLit := TCT_NumData.Create(FOwner);
      Result := FNumLit;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_NumDataSource.Write(AWriter: TXpgWriteXML);
begin
  if (FNumRef <> Nil) and FNumRef.Assigned then begin
    if xaElements in FNumRef.Assigneds then begin
      AWriter.BeginTag('c:numRef');
      FNumRef.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:numRef');
  end;
  if (FNumLit <> Nil) and FNumLit.Assigned then begin
    if xaElements in FNumLit.Assigneds then
    begin
      AWriter.BeginTag('c:numLit');
      FNumLit.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:numLit');
  end;
end;

constructor TCT_NumDataSource.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 0;
end;

destructor TCT_NumDataSource.Destroy;
begin
  if FNumRef <> Nil then 
    FNumRef.Free;
  if FNumLit <> Nil then 
    FNumLit.Free;
end;

procedure TCT_NumDataSource.Clear;
begin
  FAssigneds := [];
  if FNumRef <> Nil then 
    FreeAndNil(FNumRef);
  if FNumLit <> Nil then 
    FreeAndNil(FNumLit);
end;

procedure TCT_NumDataSource.Assign(AItem: TCT_NumDataSource);
begin
end;

procedure TCT_NumDataSource.CopyTo(AItem: TCT_NumDataSource);
begin
end;

function  TCT_NumDataSource.Create_NumRef: TCT_NumRef;
begin
  Result := TCT_NumRef.Create(FOwner);
  FNumRef := Result;
end;

function  TCT_NumDataSource.Create_NumLit: TCT_NumData;
begin
  Result := TCT_NumData.Create(FOwner);
  FNumLit := Result;
end;

{ TCT_MultiLvlStrRef }

function  TCT_MultiLvlStrRef.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FF <> '' then 
    Inc(ElemsAssigned);
  if FMultiLvlStrCache <> Nil then 
    Inc(ElemsAssigned,FMultiLvlStrCache.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_MultiLvlStrRef.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $00000103: FF := AReader.Text;
    $00000703: begin
      if FMultiLvlStrCache = Nil then 
        FMultiLvlStrCache := TCT_MultiLvlStrData.Create(FOwner);
      Result := FMultiLvlStrCache;
    end;
    $00000321: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_MultiLvlStrRef.Write(AWriter: TXpgWriteXML);
begin
  if FF <> '' then 
    AWriter.SimpleTextTag('c:f',FF);
  if (FMultiLvlStrCache <> Nil) and FMultiLvlStrCache.Assigned then 
    if xaElements in FMultiLvlStrCache.Assigneds then
    begin
      AWriter.BeginTag('c:multiLvlStrCache');
      FMultiLvlStrCache.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:multiLvlStrCache');
{$ifdef XLS_WRITE_EXTLST}
  if (FExtLst <> Nil) and FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then
    begin
      AWriter.BeginTag('c:extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:extLst');
{$endif}
end;

constructor TCT_MultiLvlStrRef.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 3;
  FAttributeCount := 0;
end;

destructor TCT_MultiLvlStrRef.Destroy;
begin
  if FMultiLvlStrCache <> Nil then 
    FMultiLvlStrCache.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_MultiLvlStrRef.Clear;
begin
  FAssigneds := [];
  FF := '';
  if FMultiLvlStrCache <> Nil then 
    FreeAndNil(FMultiLvlStrCache);
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
end;

procedure TCT_MultiLvlStrRef.Assign(AItem: TCT_MultiLvlStrRef);
begin
end;

procedure TCT_MultiLvlStrRef.CopyTo(AItem: TCT_MultiLvlStrRef);
begin
end;

function  TCT_MultiLvlStrRef.Create_MultiLvlStrCache: TCT_MultiLvlStrData;
begin
  Result := TCT_MultiLvlStrData.Create(FOwner);
  FMultiLvlStrCache := Result;
end;

function  TCT_MultiLvlStrRef.Create_ExtLst: TCT_ExtensionList;
begin
  Result := TCT_ExtensionList.Create(FOwner);
  FExtLst := Result;
end;

{ TEG_SerShared }

function  TEG_SerShared.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FIdx <> Nil then 
    Inc(ElemsAssigned,FIdx.CheckAssigned);
  if FOrder <> Nil then 
    Inc(ElemsAssigned,FOrder.CheckAssigned);
  if FTx <> Nil then 
    Inc(ElemsAssigned,FTx.CheckAssigned);
  if FSpPr <> Nil then 
    Inc(ElemsAssigned,FSpPr.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TEG_SerShared.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $000001E2: begin
      if FIdx = Nil then 
        FIdx := TCT_UnsignedInt.Create(FOwner);
      Result := FIdx;
    end;
    $000002B9: begin
      if FOrder = Nil then 
        FOrder := TCT_UnsignedInt.Create(FOwner);
      Result := FOrder;
    end;
    $00000189: begin
      if FTx = Nil then 
        FTx := TCT_SerTx.Create(FOwner);
      Result := FTx;
    end;
    $00000242: begin
      if FSpPr = Nil then 
        FSpPr := TCT_ShapeProperties.Create(FOwner);
      Result := FSpPr;
    end;
    else 
    begin
      Result := Nil;
      Exit;
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TEG_SerShared.Write(AWriter: TXpgWriteXML);
begin
  if (FIdx <> Nil) and FIdx.Assigned then 
  begin
    FIdx.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:idx');
  end;
  if (FOrder <> Nil) and FOrder.Assigned then 
  begin
    FOrder.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:order');
  end;
  if (FTx <> Nil) and FTx.Assigned then 
    if xaElements in FTx.Assigneds then
    begin
      AWriter.BeginTag('c:tx');
      FTx.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:tx');
  if (FSpPr <> Nil) and FSpPr.Assigned then 
  begin
    FSpPr.WriteAttributes(AWriter);
    if xaElements in FSpPr.Assigneds then
    begin
      AWriter.BeginTag('c:spPr');
      FSpPr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:spPr');
  end;
end;

constructor TEG_SerShared.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 4;
  FAttributeCount := 0;
end;

destructor TEG_SerShared.Destroy;
begin
  if FIdx <> Nil then 
    FIdx.Free;
  if FOrder <> Nil then 
    FOrder.Free;
  if FTx <> Nil then 
    FTx.Free;
  if FSpPr <> Nil then 
    FSpPr.Free;
end;

procedure TEG_SerShared.Clear;
begin
  FAssigneds := [];
  if FIdx <> Nil then 
    FreeAndNil(FIdx);
  if FOrder <> Nil then 
    FreeAndNil(FOrder);
  if FTx <> Nil then 
    FreeAndNil(FTx);
  if FSpPr <> Nil then 
    FreeAndNil(FSpPr);
end;

procedure TEG_SerShared.Assign(AItem: TEG_SerShared);
begin
end;

procedure TEG_SerShared.CopyTo(AItem: TEG_SerShared);
begin
end;

function  TEG_SerShared.Create_Idx: TCT_UnsignedInt;
begin
  Result := TCT_UnsignedInt.Create(FOwner);
  FIdx := Result;
end;

function  TEG_SerShared.Create_Order: TCT_UnsignedInt;
begin
  Result := TCT_UnsignedInt.Create(FOwner);
  FOrder := Result;
end;

function  TEG_SerShared.Create_Tx: TCT_SerTx;
begin
  Result := TCT_SerTx.Create(FOwner);
  FTx := Result;
end;

function  TEG_SerShared.Create_SpPr: TCT_ShapeProperties;
begin
  Result := TCT_ShapeProperties.Create(FOwner);
  FSpPr := Result;
end;

{ TCT_Marker }

function  TCT_Marker.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FSymbol <> Nil then
    Inc(ElemsAssigned,FSymbol.CheckAssigned);
  if FSize <> Nil then
    Inc(ElemsAssigned,FSize.CheckAssigned);
  if FSpPr <> Nil then
    Inc(ElemsAssigned,FSpPr.CheckAssigned);
  if FExtLst <> Nil then
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_Marker.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $00000333: begin
      if FSymbol = Nil then
        FSymbol := TCT_MarkerStyle.Create(FOwner);
      Result := FSymbol;
    end;
    $00000258: begin
      if FSize = Nil then
        FSize := TCT_MarkerSize.Create(FOwner);
      Result := FSize;
    end;
    $00000242: begin
      if FSpPr = Nil then
        FSpPr := TCT_ShapeProperties.Create(FOwner);
      Result := FSpPr;
    end;
    $00000321: begin
      if FExtLst = Nil then
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then
    Result.Assigneds := [xaRead];
end;

procedure TCT_Marker.Write(AWriter: TXpgWriteXML);
begin
  if (FSymbol <> Nil) and FSymbol.Assigned then
  begin
    FSymbol.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:symbol');
  end;
  if (FSize <> Nil) and FSize.Assigned then
  begin
    FSize.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:size');
  end;
  if (FSpPr <> Nil) and FSpPr.Assigned then
  begin
    FSpPr.WriteAttributes(AWriter);
    if xaElements in FSpPr.Assigneds then
    begin
      AWriter.BeginTag('c:spPr');
      FSpPr.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:spPr');
  end;
{$ifdef XLS_WRITE_EXTLST}
  if (FExtLst <> Nil) and FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then
    begin
      AWriter.BeginTag('c:extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:extLst');
{$endif}
end;

constructor TCT_Marker.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 4;
  FAttributeCount := 0;
end;

destructor TCT_Marker.Destroy;
begin
  if FSymbol <> Nil then
    FSymbol.Free;
  if FSize <> Nil then
    FSize.Free;
  if FSpPr <> Nil then
    FSpPr.Free;
  if FExtLst <> Nil then
    FExtLst.Free;
end;

procedure TCT_Marker.Clear;
begin
  FAssigneds := [];
  if FSymbol <> Nil then
    FreeAndNil(FSymbol);
  if FSize <> Nil then
    FreeAndNil(FSize);
  if FSpPr <> Nil then
    FreeAndNil(FSpPr);
  if FExtLst <> Nil then
    FreeAndNil(FExtLst);
end;

procedure TCT_Marker.Assign(AItem: TCT_Marker);
begin
end;

procedure TCT_Marker.CopyTo(AItem: TCT_Marker);
begin
end;

function  TCT_Marker.Create_Symbol: TCT_MarkerStyle;
begin
  Result := TCT_MarkerStyle.Create(FOwner);
  FSymbol := Result;
end;

function  TCT_Marker.Create_Size: TCT_MarkerSize;
begin
  Result := TCT_MarkerSize.Create(FOwner);
  FSize := Result;
end;

function  TCT_Marker.Create_SpPr: TCT_ShapeProperties;
begin
  Result := TCT_ShapeProperties.Create(FOwner);
  FSpPr := Result;
end;

function  TCT_Marker.Create_ExtLst: TCT_ExtensionList;
begin
  Result := TCT_ExtensionList.Create(FOwner);
  FExtLst := Result;
end;

{ TCT_DPt }

function  TCT_DPt.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FIdx <> Nil then 
    Inc(ElemsAssigned,FIdx.CheckAssigned);
  if FInvertIfNegative <> Nil then 
    Inc(ElemsAssigned,FInvertIfNegative.CheckAssigned);
  if FMarker <> Nil then 
    Inc(ElemsAssigned,FMarker.CheckAssigned);
  if FBubble3D <> Nil then 
    Inc(ElemsAssigned,FBubble3D.CheckAssigned);
  if FExplosion <> Nil then 
    Inc(ElemsAssigned,FExplosion.CheckAssigned);
  if FSpPr <> Nil then 
    Inc(ElemsAssigned,FSpPr.CheckAssigned);
  if FPictureOptions <> Nil then 
    Inc(ElemsAssigned,FPictureOptions.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_DPt.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $000001E2: begin
      if FIdx = Nil then 
        FIdx := TCT_UnsignedInt.Create(FOwner);
      Result := FIdx;
    end;
    $00000717: begin
      if FInvertIfNegative = Nil then 
        FInvertIfNegative := TCT_Boolean.Create(FOwner);
      Result := FInvertIfNegative;
    end;
    $0000031F: begin
      if FMarker = Nil then 
        FMarker := TCT_Marker.Create(FOwner);
      Result := FMarker;
    end;
    $00000380: begin
      if FBubble3D = Nil then 
        FBubble3D := TCT_Boolean.Create(FOwner);
      Result := FBubble3D;
    end;
    $0000047E: begin
      if FExplosion = Nil then 
        FExplosion := TCT_UnsignedInt.Create(FOwner);
      Result := FExplosion;
    end;
    $00000242: begin
      if FSpPr = Nil then 
        FSpPr := TCT_ShapeProperties.Create(FOwner);
      Result := FSpPr;
    end;
    $00000685: begin
      if FPictureOptions = Nil then 
        FPictureOptions := TCT_PictureOptions.Create(FOwner);
      Result := FPictureOptions;
    end;
    $00000321: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_DPt.Write(AWriter: TXpgWriteXML);
begin
  if (FIdx <> Nil) and FIdx.Assigned then 
  begin
    FIdx.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:idx');
  end
  else
    AWriter.SimpleTag('c:idx');
  if (FInvertIfNegative <> Nil) and FInvertIfNegative.Assigned then
  begin
    FInvertIfNegative.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:invertIfNegative');
  end;
  if (FMarker <> Nil) and FMarker.Assigned then
    if xaElements in FMarker.Assigneds then
    begin
      AWriter.BeginTag('c:marker');
      FMarker.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:marker');
  if (FBubble3D <> Nil) and FBubble3D.Assigned then
  begin
    FBubble3D.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:bubble3D');
  end;
  if (FExplosion <> Nil) and FExplosion.Assigned then
  begin
    FExplosion.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:explosion');
  end;
  if (FSpPr <> Nil) and FSpPr.Assigned then
  begin
    FSpPr.WriteAttributes(AWriter);
    if xaElements in FSpPr.Assigneds then
    begin
      AWriter.BeginTag('c:spPr');
      FSpPr.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:spPr');
  end;
  if (FPictureOptions <> Nil) and FPictureOptions.Assigned then
    if xaElements in FPictureOptions.Assigneds then
    begin
      AWriter.BeginTag('c:pictureOptions');
      FPictureOptions.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:pictureOptions');
{$ifdef XLS_WRITE_EXTLST}
  if (FExtLst <> Nil) and FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then
    begin
      AWriter.BeginTag('c:extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:extLst');
{$endif}
end;

constructor TCT_DPt.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 8;
  FAttributeCount := 0;
end;

destructor TCT_DPt.Destroy;
begin
  if FIdx <> Nil then 
    FIdx.Free;
  if FInvertIfNegative <> Nil then 
    FInvertIfNegative.Free;
  if FMarker <> Nil then 
    FMarker.Free;
  if FBubble3D <> Nil then 
    FBubble3D.Free;
  if FExplosion <> Nil then 
    FExplosion.Free;
  if FSpPr <> Nil then 
    FSpPr.Free;
  if FPictureOptions <> Nil then 
    FPictureOptions.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_DPt.Clear;
begin
  FAssigneds := [];
  if FIdx <> Nil then 
    FreeAndNil(FIdx);
  if FInvertIfNegative <> Nil then 
    FreeAndNil(FInvertIfNegative);
  if FMarker <> Nil then 
    FreeAndNil(FMarker);
  if FBubble3D <> Nil then 
    FreeAndNil(FBubble3D);
  if FExplosion <> Nil then 
    FreeAndNil(FExplosion);
  if FSpPr <> Nil then 
    FreeAndNil(FSpPr);
  if FPictureOptions <> Nil then 
    FreeAndNil(FPictureOptions);
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
end;

procedure TCT_DPt.Assign(AItem: TCT_DPt);
begin
end;

procedure TCT_DPt.CopyTo(AItem: TCT_DPt);
begin
end;

function  TCT_DPt.Create_Idx: TCT_UnsignedInt;
begin
  Result := TCT_UnsignedInt.Create(FOwner);
  FIdx := Result;
end;

function  TCT_DPt.Create_InvertIfNegative: TCT_Boolean;
begin
  Result := TCT_Boolean.Create(FOwner);
  FInvertIfNegative := Result;
end;

function  TCT_DPt.Create_Marker: TCT_Marker;
begin
  Result := TCT_Marker.Create(FOwner);
  FMarker := Result;
end;

function  TCT_DPt.Create_Bubble3D: TCT_Boolean;
begin
  Result := TCT_Boolean.Create(FOwner);
  FBubble3D := Result;
end;

function  TCT_DPt.Create_Explosion: TCT_UnsignedInt;
begin
  Result := TCT_UnsignedInt.Create(FOwner);
  FExplosion := Result;
end;

function  TCT_DPt.Create_SpPr: TCT_ShapeProperties;
begin
  Result := TCT_ShapeProperties.Create(FOwner);
  FSpPr := Result;
end;

function  TCT_DPt.Create_PictureOptions: TCT_PictureOptions;
begin
  Result := TCT_PictureOptions.Create(FOwner);
  FPictureOptions := Result;
end;

function  TCT_DPt.Create_ExtLst: TCT_ExtensionList;
begin
  Result := TCT_ExtensionList.Create(FOwner);
  FExtLst := Result;
end;

{ TCT_DPtXpgList }

function  TCT_DPtXpgList.GetItems(Index: integer): TCT_DPt;
begin
  Result := TCT_DPt(inherited Items[Index]);
end;

function  TCT_DPtXpgList.Add: TCT_DPt;
begin
  Result := TCT_DPt.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_DPtXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_DPtXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
    if xaElements in Items[i].Assigneds then
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
end;

procedure TCT_DPtXpgList.Assign(AItem: TCT_DPtXpgList);
begin
end;

procedure TCT_DPtXpgList.CopyTo(AItem: TCT_DPtXpgList);
begin
end;

{ TCT_DLbls }

function  TCT_DLbls.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FDLblXpgList <> Nil then 
    Inc(ElemsAssigned,FDLblXpgList.CheckAssigned);
  if FDelete <> Nil then 
    Inc(ElemsAssigned,FDelete.CheckAssigned);
  Inc(ElemsAssigned,FGroup_DLbls.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_DLbls.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $0000021B: begin
      if FDLblXpgList = Nil then 
        FDLblXpgList := TCT_DLblXpgList.Create(FOwner);
      Result := FDLblXpgList.Add;
    end;
    $00000310: begin
      if FDelete = Nil then 
        FDelete := TCT_Boolean.Create(FOwner);
      Result := FDelete;
    end;
    $00000321: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
    begin
      Result := FGroup_DLbls.HandleElement(AReader);
      if Result = Nil then 
        FOwner.Errors.Error(xemUnknownElement,AReader.QName);
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_DLbls.Write(AWriter: TXpgWriteXML);
begin
  if FDLblXpgList <> Nil then
    FDLblXpgList.Write(AWriter,'c:dLbl');
  if (FDelete <> Nil) and FDelete.Assigned then 
  begin
    FDelete.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:delete');
  end;
  FGroup_DLbls.Write(AWriter);
{$ifdef XLS_WRITE_EXTLST}
  if (FExtLst <> Nil) and FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then
    begin
      AWriter.BeginTag('c:extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:extLst');
{$endif}
end;

constructor TCT_DLbls.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 4;
  FAttributeCount := 0;
  FGroup_DLbls := TGroup_DLbls.Create(FOwner);
end;

destructor TCT_DLbls.Destroy;
begin
  if FDLblXpgList <> Nil then 
    FDLblXpgList.Free;
  if FDelete <> Nil then 
    FDelete.Free;
  FGroup_DLbls.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_DLbls.Clear;
begin
  FAssigneds := [];
  if FDLblXpgList <> Nil then 
    FreeAndNil(FDLblXpgList);
  if FDelete <> Nil then 
    FreeAndNil(FDelete);
  FGroup_DLbls.Clear;
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
end;

procedure TCT_DLbls.Assign(AItem: TCT_DLbls);
begin
end;

procedure TCT_DLbls.CopyTo(AItem: TCT_DLbls);
begin
end;

function  TCT_DLbls.Create_DLblXpgList: TCT_DLblXpgList;
begin
  Result := TCT_DLblXpgList.Create(FOwner);
  FDLblXpgList := Result;
end;

function  TCT_DLbls.Create_Delete: TCT_Boolean;
begin
  Result := TCT_Boolean.Create(FOwner);
  FDelete := Result;
end;

function  TCT_DLbls.Create_ExtLst: TCT_ExtensionList;
begin
  Result := TCT_ExtensionList.Create(FOwner);
  FExtLst := Result;
end;

{ TCT_Trendline }

function  TCT_Trendline.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FName <> '' then 
    Inc(ElemsAssigned);
  if FSpPr <> Nil then 
    Inc(ElemsAssigned,FSpPr.CheckAssigned);
  if FTrendlineType <> Nil then 
    Inc(ElemsAssigned,FTrendlineType.CheckAssigned);
  if FOrder <> Nil then 
    Inc(ElemsAssigned,FOrder.CheckAssigned);
  if FPeriod <> Nil then 
    Inc(ElemsAssigned,FPeriod.CheckAssigned);
  if FForward <> Nil then 
    Inc(ElemsAssigned,FForward.CheckAssigned);
  if FBackward <> Nil then 
    Inc(ElemsAssigned,FBackward.CheckAssigned);
  if FIntercept <> Nil then 
    Inc(ElemsAssigned,FIntercept.CheckAssigned);
  if FDispRSqr <> Nil then 
    Inc(ElemsAssigned,FDispRSqr.CheckAssigned);
  if FDispEq <> Nil then 
    Inc(ElemsAssigned,FDispEq.CheckAssigned);
  if FTrendlineLbl <> Nil then 
    Inc(ElemsAssigned,FTrendlineLbl.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_Trendline.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $0000023E: FName := AReader.Text;
    $00000242: begin
      if FSpPr = Nil then 
        FSpPr := TCT_ShapeProperties.Create(FOwner);
      Result := FSpPr;
    end;
    $00000604: begin
      if FTrendlineType = Nil then 
        FTrendlineType := TCT_TrendlineType.Create(FOwner);
      Result := FTrendlineType;
    end;
    $000002B9: begin
      if FOrder = Nil then 
        FOrder := TCT_Order.Create(FOwner);
      Result := FOrder;
    end;
    $00000320: begin
      if FPeriod = Nil then 
        FPeriod := TCT_Period.Create(FOwner);
      Result := FPeriod;
    end;
    $00000392: begin
      if FForward = Nil then 
        FForward := TCT_Double.Create(FOwner);
      Result := FForward;
    end;
    $000003DC: begin
      if FBackward = Nil then 
        FBackward := TCT_Double.Create(FOwner);
      Result := FBackward;
    end;
    $0000046B: begin
      if FIntercept = Nil then 
        FIntercept := TCT_Double.Create(FOwner);
      Result := FIntercept;
    end;
    $000003D5: begin
      if FDispRSqr = Nil then 
        FDispRSqr := TCT_Boolean.Create(FOwner);
      Result := FDispRSqr;
    end;
    $00000303: begin
      if FDispEq = Nil then 
        FDispEq := TCT_Boolean.Create(FOwner);
      Result := FDispEq;
    end;
    $0000057C: begin
      if FTrendlineLbl = Nil then 
        FTrendlineLbl := TCT_TrendlineLbl.Create(FOwner);
      Result := FTrendlineLbl;
    end;
    $00000321: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_Trendline.Write(AWriter: TXpgWriteXML);
begin
  if FName <> '' then 
    AWriter.SimpleTextTag('c:name',FName);
  if (FSpPr <> Nil) and FSpPr.Assigned then 
  begin
    FSpPr.WriteAttributes(AWriter);
    if xaElements in FSpPr.Assigneds then
    begin
      AWriter.BeginTag('c:spPr');
      FSpPr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:spPr');
  end;

  // Must be written
  FTrendlineType.WriteAttributes(AWriter);
  AWriter.SimpleTag('c:trendlineType');

  if (FOrder <> Nil) and FOrder.Assigned then
  begin
    FOrder.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:order');
  end;
  if (FPeriod <> Nil) and FPeriod.Assigned then 
  begin
    FPeriod.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:period');
  end;
  if (FForward <> Nil) and FForward.Assigned then 
  begin
    FForward.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:forward');
  end;
  if (FBackward <> Nil) and FBackward.Assigned then 
  begin
    FBackward.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:backward');
  end;
  if (FIntercept <> Nil) and FIntercept.Assigned then 
  begin
    FIntercept.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:intercept');
  end;
  if (FDispRSqr <> Nil) and FDispRSqr.Assigned then 
  begin
    FDispRSqr.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:dispRSqr');
  end;
  if (FDispEq <> Nil) and FDispEq.Assigned then 
  begin
    FDispEq.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:dispEq');
  end;
  if (FTrendlineLbl <> Nil) and FTrendlineLbl.Assigned then 
    if xaElements in FTrendlineLbl.Assigneds then
    begin
      AWriter.BeginTag('c:trendlineLbl');
      FTrendlineLbl.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:trendlineLbl');
{$ifdef XLS_WRITE_EXTLST}
  if (FExtLst <> Nil) and FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then
    begin
      AWriter.BeginTag('c:extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:extLst');
{$endif}
end;

constructor TCT_Trendline.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 12;
  FAttributeCount := 0;
end;

destructor TCT_Trendline.Destroy;
begin
  if FSpPr <> Nil then 
    FSpPr.Free;
  if FTrendlineType <> Nil then 
    FTrendlineType.Free;
  if FOrder <> Nil then 
    FOrder.Free;
  if FPeriod <> Nil then 
    FPeriod.Free;
  if FForward <> Nil then 
    FForward.Free;
  if FBackward <> Nil then 
    FBackward.Free;
  if FIntercept <> Nil then 
    FIntercept.Free;
  if FDispRSqr <> Nil then 
    FDispRSqr.Free;
  if FDispEq <> Nil then 
    FDispEq.Free;
  if FTrendlineLbl <> Nil then 
    FTrendlineLbl.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_Trendline.Clear;
begin
  FAssigneds := [];
  FName := '';
  if FSpPr <> Nil then 
    FreeAndNil(FSpPr);
  if FTrendlineType <> Nil then 
    FreeAndNil(FTrendlineType);
  if FOrder <> Nil then 
    FreeAndNil(FOrder);
  if FPeriod <> Nil then 
    FreeAndNil(FPeriod);
  if FForward <> Nil then 
    FreeAndNil(FForward);
  if FBackward <> Nil then 
    FreeAndNil(FBackward);
  if FIntercept <> Nil then 
    FreeAndNil(FIntercept);
  if FDispRSqr <> Nil then 
    FreeAndNil(FDispRSqr);
  if FDispEq <> Nil then 
    FreeAndNil(FDispEq);
  if FTrendlineLbl <> Nil then 
    FreeAndNil(FTrendlineLbl);
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
end;

procedure TCT_Trendline.Assign(AItem: TCT_Trendline);
begin
end;

procedure TCT_Trendline.CopyTo(AItem: TCT_Trendline);
begin
end;

function  TCT_Trendline.Create_SpPr: TCT_ShapeProperties;
begin
  Result := TCT_ShapeProperties.Create(FOwner);
  FSpPr := Result;
end;

function  TCT_Trendline.Create_TrendlineType: TCT_TrendlineType;
begin
  Result := TCT_TrendlineType.Create(FOwner);
  FTrendlineType := Result;
end;

function  TCT_Trendline.Create_Order: TCT_Order;
begin
  Result := TCT_Order.Create(FOwner);
  FOrder := Result;
end;

function  TCT_Trendline.Create_Period: TCT_Period;
begin
  Result := TCT_Period.Create(FOwner);
  FPeriod := Result;
end;

function  TCT_Trendline.Create_Forward: TCT_Double;
begin
  Result := TCT_Double.Create(FOwner);
  FForward := Result;
end;

function  TCT_Trendline.Create_Backward: TCT_Double;
begin
  Result := TCT_Double.Create(FOwner);
  FBackward := Result;
end;

function  TCT_Trendline.Create_Intercept: TCT_Double;
begin
  Result := TCT_Double.Create(FOwner);
  FIntercept := Result;
end;

function  TCT_Trendline.Create_DispRSqr: TCT_Boolean;
begin
  Result := TCT_Boolean.Create(FOwner);
  FDispRSqr := Result;
end;

function  TCT_Trendline.Create_DispEq: TCT_Boolean;
begin
  Result := TCT_Boolean.Create(FOwner);
  FDispEq := Result;
end;

function  TCT_Trendline.Create_TrendlineLbl: TCT_TrendlineLbl;
begin
  Result := TCT_TrendlineLbl.Create(FOwner);
  FTrendlineLbl := Result;
end;

function  TCT_Trendline.Create_ExtLst: TCT_ExtensionList;
begin
  Result := TCT_ExtensionList.Create(FOwner);
  FExtLst := Result;
end;

{ TCT_TrendlineXpgList }

function  TCT_TrendlineXpgList.GetItems(Index: integer): TCT_Trendline;
begin
  Result := TCT_Trendline(inherited Items[Index]);
end;

function  TCT_TrendlineXpgList.Add: TCT_Trendline;
begin
  Result := TCT_Trendline.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_TrendlineXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_TrendlineXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
    if xaElements in Items[i].Assigneds then
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
end;

procedure TCT_TrendlineXpgList.Assign(AItem: TCT_TrendlineXpgList);
begin
end;

procedure TCT_TrendlineXpgList.CopyTo(AItem: TCT_TrendlineXpgList);
begin
end;

{ TCT_ErrBars }

function  TCT_ErrBars.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FErrDir <> Nil then 
    Inc(ElemsAssigned,FErrDir.CheckAssigned);
  if FErrBarType <> Nil then 
    Inc(ElemsAssigned,FErrBarType.CheckAssigned);
  if FErrValType <> Nil then 
    Inc(ElemsAssigned,FErrValType.CheckAssigned);
  if FNoEndCap <> Nil then 
    Inc(ElemsAssigned,FNoEndCap.CheckAssigned);
  if FPlus <> Nil then 
    Inc(ElemsAssigned,FPlus.CheckAssigned);
  if FMinus <> Nil then 
    Inc(ElemsAssigned,FMinus.CheckAssigned);
  if FVal <> Nil then 
    Inc(ElemsAssigned,FVal.CheckAssigned);
  if FSpPr <> Nil then 
    Inc(ElemsAssigned,FSpPr.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_ErrBars.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $00000305: begin
      if FErrDir = Nil then 
        FErrDir := TCT_ErrDir.Create(FOwner);
      Result := FErrDir;
    end;
    $0000049D: begin
      if FErrBarType = Nil then 
        FErrBarType := TCT_ErrBarType.Create(FOwner);
      Result := FErrBarType;
    end;
    $000004AB: begin
      if FErrValType = Nil then 
        FErrValType := TCT_ErrValType.Create(FOwner);
      Result := FErrValType;
    end;
    $000003A5: begin
      if FNoEndCap = Nil then 
        FNoEndCap := TCT_Boolean.Create(FOwner);
      Result := FNoEndCap;
    end;
    $00000261: begin
      if FPlus = Nil then 
        FPlus := TCT_NumDataSource.Create(FOwner);
      Result := FPlus;
    end;
    $000002C9: begin
      if FMinus = Nil then 
        FMinus := TCT_NumDataSource.Create(FOwner);
      Result := FMinus;
    end;
    $000001E0: begin
      if FVal = Nil then 
        FVal := TCT_Double.Create(FOwner);
      Result := FVal;
    end;
    $00000242: begin
      if FSpPr = Nil then 
        FSpPr := TCT_ShapeProperties.Create(FOwner);
      Result := FSpPr;
    end;
    $00000321: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_ErrBars.Write(AWriter: TXpgWriteXML);
begin
  if (FErrDir <> Nil) and FErrDir.Assigned then 
  begin
    FErrDir.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:errDir');
  end;
  if (FErrBarType <> Nil) and FErrBarType.Assigned then 
  begin
    FErrBarType.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:errBarType');
  end
  else 
    AWriter.SimpleTag('c:errBarType');
  if (FErrValType <> Nil) and FErrValType.Assigned then 
  begin
    FErrValType.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:errValType');
  end
  else 
    AWriter.SimpleTag('c:errValType');
  if (FNoEndCap <> Nil) and FNoEndCap.Assigned then 
  begin
    FNoEndCap.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:noEndCap');
  end;
  if (FPlus <> Nil) and FPlus.Assigned then 
    if xaElements in FPlus.Assigneds then
    begin
      AWriter.BeginTag('c:plus');
      FPlus.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:plus');
  if (FMinus <> Nil) and FMinus.Assigned then 
    if xaElements in FMinus.Assigneds then
    begin
      AWriter.BeginTag('c:minus');
      FMinus.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:minus');
  if (FVal <> Nil) and FVal.Assigned then 
  begin
    FVal.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:val');
  end;
  if (FSpPr <> Nil) and FSpPr.Assigned then 
  begin
    FSpPr.WriteAttributes(AWriter);
    if xaElements in FSpPr.Assigneds then
    begin
      AWriter.BeginTag('c:spPr');
      FSpPr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:spPr');
  end;
{$ifdef XLS_WRITE_EXTLST}
  if (FExtLst <> Nil) and FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then
    begin
      AWriter.BeginTag('c:extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:extLst');
{$endif}
end;

constructor TCT_ErrBars.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 9;
  FAttributeCount := 0;
end;

destructor TCT_ErrBars.Destroy;
begin
  if FErrDir <> Nil then 
    FErrDir.Free;
  if FErrBarType <> Nil then 
    FErrBarType.Free;
  if FErrValType <> Nil then 
    FErrValType.Free;
  if FNoEndCap <> Nil then 
    FNoEndCap.Free;
  if FPlus <> Nil then 
    FPlus.Free;
  if FMinus <> Nil then 
    FMinus.Free;
  if FVal <> Nil then 
    FVal.Free;
  if FSpPr <> Nil then 
    FSpPr.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_ErrBars.Clear;
begin
  FAssigneds := [];
  if FErrDir <> Nil then 
    FreeAndNil(FErrDir);
  if FErrBarType <> Nil then 
    FreeAndNil(FErrBarType);
  if FErrValType <> Nil then 
    FreeAndNil(FErrValType);
  if FNoEndCap <> Nil then 
    FreeAndNil(FNoEndCap);
  if FPlus <> Nil then 
    FreeAndNil(FPlus);
  if FMinus <> Nil then 
    FreeAndNil(FMinus);
  if FVal <> Nil then 
    FreeAndNil(FVal);
  if FSpPr <> Nil then 
    FreeAndNil(FSpPr);
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
end;

procedure TCT_ErrBars.Assign(AItem: TCT_ErrBars);
begin
end;

procedure TCT_ErrBars.CopyTo(AItem: TCT_ErrBars);
begin
end;

function  TCT_ErrBars.Create_ErrDir: TCT_ErrDir;
begin
  Result := TCT_ErrDir.Create(FOwner);
  FErrDir := Result;
end;

function  TCT_ErrBars.Create_ErrBarType: TCT_ErrBarType;
begin
  Result := TCT_ErrBarType.Create(FOwner);
  FErrBarType := Result;
end;

function  TCT_ErrBars.Create_ErrValType: TCT_ErrValType;
begin
  Result := TCT_ErrValType.Create(FOwner);
  FErrValType := Result;
end;

function  TCT_ErrBars.Create_NoEndCap: TCT_Boolean;
begin
  Result := TCT_Boolean.Create(FOwner);
  FNoEndCap := Result;
end;

function  TCT_ErrBars.Create_Plus: TCT_NumDataSource;
begin
  Result := TCT_NumDataSource.Create(FOwner);
  FPlus := Result;
end;

function  TCT_ErrBars.Create_Minus: TCT_NumDataSource;
begin
  Result := TCT_NumDataSource.Create(FOwner);
  FMinus := Result;
end;

function  TCT_ErrBars.Create_Val: TCT_Double;
begin
  Result := TCT_Double.Create(FOwner);
  FVal := Result;
end;

function  TCT_ErrBars.Create_SpPr: TCT_ShapeProperties;
begin
  Result := TCT_ShapeProperties.Create(FOwner);
  FSpPr := Result;
end;

function  TCT_ErrBars.Create_ExtLst: TCT_ExtensionList;
begin
  Result := TCT_ExtensionList.Create(FOwner);
  FExtLst := Result;
end;

{ TCT_ErrBarsXpgList }

function  TCT_ErrBarsXpgList.GetItems(Index: integer): TCT_ErrBars;
begin
  Result := TCT_ErrBars(inherited Items[Index]);
end;

function  TCT_ErrBarsXpgList.Add: TCT_ErrBars;
begin
  Result := TCT_ErrBars.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_ErrBarsXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_ErrBarsXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
    if xaElements in Items[i].Assigneds then
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
end;

procedure TCT_ErrBarsXpgList.Assign(AItem: TCT_ErrBarsXpgList);
begin
end;

procedure TCT_ErrBarsXpgList.CopyTo(AItem: TCT_ErrBarsXpgList);
begin
end;

{ TCT_AxDataSource }

function  TCT_AxDataSource.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FMultiLvlStrRef <> Nil then 
    Inc(ElemsAssigned,FMultiLvlStrRef.CheckAssigned);
  if FNumRef <> Nil then 
    Inc(ElemsAssigned,FNumRef.CheckAssigned);
  if FNumLit <> Nil then 
    Inc(ElemsAssigned,FNumLit.CheckAssigned);
  if FStrRef <> Nil then 
    Inc(ElemsAssigned,FStrRef.CheckAssigned);
  if FStrLit <> Nil then 
    Inc(ElemsAssigned,FStrLit.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_AxDataSource.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $0000064C: begin
      if FMultiLvlStrRef = Nil then 
        FMultiLvlStrRef := TCT_MultiLvlStrRef.Create(FOwner);
      Result := FMultiLvlStrRef;
    end;
    $0000030A: begin
      if FNumRef = Nil then 
        FNumRef := TCT_NumRef.Create(FOwner);
      Result := FNumRef;
    end;
    $00000316: begin
      if FNumLit = Nil then 
        FNumLit := TCT_NumData.Create(FOwner);
      Result := FNumLit;
    end;
    $00000313: begin
      if FStrRef = Nil then 
        FStrRef := TCT_StrRef.Create(FOwner);
      Result := FStrRef;
    end;
    $0000031F: begin
      if FStrLit = Nil then 
        FStrLit := TCT_StrData.Create(FOwner);
      Result := FStrLit;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_AxDataSource.Write(AWriter: TXpgWriteXML);
begin
  if (FMultiLvlStrRef <> Nil) and FMultiLvlStrRef.Assigned then 
    if xaElements in FMultiLvlStrRef.Assigneds then
    begin
      AWriter.BeginTag('c:multiLvlStrRef');
      FMultiLvlStrRef.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:multiLvlStrRef');
//  else
//    AWriter.SimpleTag('c:multiLvlStrRef');
  if (FNumRef <> Nil) and FNumRef.Assigned then
    if xaElements in FNumRef.Assigneds then
    begin
      AWriter.BeginTag('c:numRef');
      FNumRef.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:numRef');
  if (FNumLit <> Nil) and FNumLit.Assigned then
    if xaElements in FNumLit.Assigneds then
    begin
      AWriter.BeginTag('c:numLit');
      FNumLit.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:numLit');
  if (FStrRef <> Nil) and FStrRef.Assigned then
    if xaElements in FStrRef.Assigneds then
    begin
      AWriter.BeginTag('c:strRef');
      FStrRef.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:strRef');
  if (FStrLit <> Nil) and FStrLit.Assigned then
    if xaElements in FStrLit.Assigneds then
    begin
      AWriter.BeginTag('c:strLit');
      FStrLit.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:strLit');
end;

constructor TCT_AxDataSource.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 5;
  FAttributeCount := 0;
end;

destructor TCT_AxDataSource.Destroy;
begin
  if FMultiLvlStrRef <> Nil then 
    FMultiLvlStrRef.Free;
  if FNumRef <> Nil then 
    FNumRef.Free;
  if FNumLit <> Nil then 
    FNumLit.Free;
  if FStrRef <> Nil then 
    FStrRef.Free;
  if FStrLit <> Nil then 
    FStrLit.Free;
end;

procedure TCT_AxDataSource.Clear;
begin
  FAssigneds := [];
  if FMultiLvlStrRef <> Nil then 
    FreeAndNil(FMultiLvlStrRef);
  if FNumRef <> Nil then 
    FreeAndNil(FNumRef);
  if FNumLit <> Nil then 
    FreeAndNil(FNumLit);
  if FStrRef <> Nil then 
    FreeAndNil(FStrRef);
  if FStrLit <> Nil then 
    FreeAndNil(FStrLit);
end;

procedure TCT_AxDataSource.Assign(AItem: TCT_AxDataSource);
begin
end;

procedure TCT_AxDataSource.CopyTo(AItem: TCT_AxDataSource);
begin
end;

function  TCT_AxDataSource.Create_MultiLvlStrRef: TCT_MultiLvlStrRef;
begin
  Result := TCT_MultiLvlStrRef.Create(FOwner);
  FMultiLvlStrRef := Result;
end;

function  TCT_AxDataSource.Create_NumRef: TCT_NumRef;
begin
  Result := TCT_NumRef.Create(FOwner);
  FNumRef := Result;
end;

function  TCT_AxDataSource.Create_NumLit: TCT_NumData;
begin
  Result := TCT_NumData.Create(FOwner);
  FNumLit := Result;
end;

function  TCT_AxDataSource.Create_StrRef: TCT_StrRef;
begin
  Result := TCT_StrRef.Create(FOwner);
  FStrRef := Result;
end;

function  TCT_AxDataSource.Create_StrLit: TCT_StrData;
begin
  Result := TCT_StrData.Create(FOwner);
  FStrLit := Result;
end;

{ TCT_BandFmt }

function  TCT_BandFmt.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FIdx <> Nil then 
    Inc(ElemsAssigned,FIdx.CheckAssigned);
  if FSpPr <> Nil then 
    Inc(ElemsAssigned,FSpPr.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_BandFmt.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $000001E2: begin
      if FIdx = Nil then 
        FIdx := TCT_UnsignedInt.Create(FOwner);
      Result := FIdx;
    end;
    $00000242: begin
      if FSpPr = Nil then 
        FSpPr := TCT_ShapeProperties.Create(FOwner);
      Result := FSpPr;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_BandFmt.Write(AWriter: TXpgWriteXML);
begin
  if (FIdx <> Nil) and FIdx.Assigned then 
  begin
    FIdx.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:idx');
  end
  else 
    AWriter.SimpleTag('c:idx');
  if (FSpPr <> Nil) and FSpPr.Assigned then 
  begin
    FSpPr.WriteAttributes(AWriter);
    if xaElements in FSpPr.Assigneds then
    begin
      AWriter.BeginTag('c:spPr');
      FSpPr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:spPr');
  end;
end;

constructor TCT_BandFmt.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 0;
end;

destructor TCT_BandFmt.Destroy;
begin
  if FIdx <> Nil then 
    FIdx.Free;
  if FSpPr <> Nil then 
    FSpPr.Free;
end;

procedure TCT_BandFmt.Clear;
begin
  FAssigneds := [];
  if FIdx <> Nil then 
    FreeAndNil(FIdx);
  if FSpPr <> Nil then 
    FreeAndNil(FSpPr);
end;

procedure TCT_BandFmt.Assign(AItem: TCT_BandFmt);
begin
end;

procedure TCT_BandFmt.CopyTo(AItem: TCT_BandFmt);
begin
end;

function  TCT_BandFmt.Create_Idx: TCT_UnsignedInt;
begin
  Result := TCT_UnsignedInt.Create(FOwner);
  FIdx := Result;
end;

function  TCT_BandFmt.Create_SpPr: TCT_ShapeProperties;
begin
  Result := TCT_ShapeProperties.Create(FOwner);
  FSpPr := Result;
end;

{ TCT_BandFmtXpgList }

function  TCT_BandFmtXpgList.GetItems(Index: integer): TCT_BandFmt;
begin
  Result := TCT_BandFmt(inherited Items[Index]);
end;

function  TCT_BandFmtXpgList.Add: TCT_BandFmt;
begin
  Result := TCT_BandFmt.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_BandFmtXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_BandFmtXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
    if xaElements in Items[i].Assigneds then
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
end;

procedure TCT_BandFmtXpgList.Assign(AItem: TCT_BandFmtXpgList);
begin
end;

procedure TCT_BandFmtXpgList.CopyTo(AItem: TCT_BandFmtXpgList);
begin
end;

{ TCT_LogBase }

function  TCT_LogBase.CheckAssigned: integer;
begin
  FAssigneds := [xaAttributes];
  Result := 1;
end;

procedure TCT_LogBase.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_LogBase.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('val',XmlFloatToStr(FVal));
end;

procedure TCT_LogBase.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then 
    FVal := XmlStrToFloatDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_LogBase.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := NaN;
end;

destructor TCT_LogBase.Destroy;
begin
end;

procedure TCT_LogBase.Clear;
begin
  FAssigneds := [];
  FVal := NaN;
end;

procedure TCT_LogBase.Assign(AItem: TCT_LogBase);
begin
end;

procedure TCT_LogBase.CopyTo(AItem: TCT_LogBase);
begin
end;

{ TCT_Orientation }

function  TCT_Orientation.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FVal <> stoMinMax then
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then
    FAssigneds := [xaAttributes];
end;

procedure TCT_Orientation.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_Orientation.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FVal <> stoMinMax then
    AWriter.AddAttribute('val',StrTST_Orientation[Integer(FVal)]);
end;

procedure TCT_Orientation.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then
    FVal := TST_Orientation(StrToEnum('sto' + AAttributes.Values[0]))
  else
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_Orientation.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := stoMinMax;
end;

destructor TCT_Orientation.Destroy;
begin
end;

procedure TCT_Orientation.Clear;
begin
  FAssigneds := [];
  FVal := stoMinMax;
end;

procedure TCT_Orientation.Assign(AItem: TCT_Orientation);
begin
end;

procedure TCT_Orientation.CopyTo(AItem: TCT_Orientation);
begin
end;

{ TCT_Grouping }

function  TCT_Grouping.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FVal <> stgStandard then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_Grouping.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_Grouping.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FVal <> stgStandard then 
    AWriter.AddAttribute('val',StrTST_Grouping[Integer(FVal)]);
end;

procedure TCT_Grouping.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then 
    FVal := TST_Grouping(StrToEnum('stg' + AAttributes.Values[0]))
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_Grouping.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := stgStandard;
end;

destructor TCT_Grouping.Destroy;
begin
end;

procedure TCT_Grouping.Clear;
begin
  FAssigneds := [];
  FVal := stgStandard;
end;

procedure TCT_Grouping.Assign(AItem: TCT_Grouping);
begin
end;

procedure TCT_Grouping.CopyTo(AItem: TCT_Grouping);
begin
end;

{ TCT_AreaSer }

function  TCT_AreaSer.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FEG_SerShared.CheckAssigned);
  if FPictureOptions <> Nil then 
    Inc(ElemsAssigned,FPictureOptions.CheckAssigned);
  if FDPtXpgList <> Nil then 
    Inc(ElemsAssigned,FDPtXpgList.CheckAssigned);
  if FDLbls <> Nil then 
    Inc(ElemsAssigned,FDLbls.CheckAssigned);
  if FTrendlineXpgList <> Nil then 
    Inc(ElemsAssigned,FTrendlineXpgList.CheckAssigned);
  if FErrBarsXpgList <> Nil then 
    Inc(ElemsAssigned,FErrBarsXpgList.CheckAssigned);
  if FCat <> Nil then 
    Inc(ElemsAssigned,FCat.CheckAssigned);
  if FVal <> Nil then 
    Inc(ElemsAssigned,FVal.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_AreaSer.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $00000685: begin
      if FPictureOptions = Nil then 
        FPictureOptions := TCT_PictureOptions.Create(FOwner);
      Result := FPictureOptions;
    end;
    $000001C5: begin
      if FDPtXpgList = Nil then 
        FDPtXpgList := TCT_DPtXpgList.Create(FOwner);
      Result := FDPtXpgList.Add;
    end;
    $0000028E: begin
      if FDLbls = Nil then 
        FDLbls := TCT_DLbls.Create(FOwner);
      Result := FDLbls;
    end;
    $00000462: begin
      if FTrendlineXpgList = Nil then 
        FTrendlineXpgList := TCT_TrendlineXpgList.Create(FOwner);
      Result := FTrendlineXpgList.Add;
    end;
    $0000036E: begin
      if FErrBarsXpgList = Nil then 
        FErrBarsXpgList := TCT_ErrBarsXpgList.Create(FOwner);
      Result := FErrBarsXpgList.Add;
    end;
    $000001D5: begin
      if FCat = Nil then 
        FCat := TCT_AxDataSource.Create(FOwner);
      Result := FCat;
    end;
    $000001E0: begin
      if FVal = Nil then 
        FVal := TCT_NumDataSource.Create(FOwner);
      Result := FVal;
    end;
    $00000321: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
    begin
      Result := FEG_SerShared.HandleElement(AReader);
      if Result = Nil then 
        FOwner.Errors.Error(xemUnknownElement,AReader.QName);
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_AreaSer.Write(AWriter: TXpgWriteXML);
begin
  FEG_SerShared.Write(AWriter);
  if (FPictureOptions <> Nil) and FPictureOptions.Assigned then 
    if xaElements in FPictureOptions.Assigneds then
    begin
      AWriter.BeginTag('c:pictureOptions');
      FPictureOptions.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:pictureOptions');
  if FDPtXpgList <> Nil then 
    FDPtXpgList.Write(AWriter,'c:dPt');
  if (FDLbls <> Nil) and FDLbls.Assigned then 
    if xaElements in FDLbls.Assigneds then
    begin
      AWriter.BeginTag('c:dLbls');
      FDLbls.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:dLbls');
  if FTrendlineXpgList <> Nil then 
    FTrendlineXpgList.Write(AWriter,'c:trendline');
  if FErrBarsXpgList <> Nil then 
    FErrBarsXpgList.Write(AWriter,'c:errBars');
  if (FCat <> Nil) and FCat.Assigned then 
    if xaElements in FCat.Assigneds then
    begin
      AWriter.BeginTag('c:cat');
      FCat.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:cat');
  if (FVal <> Nil) and FVal.Assigned then 
    if xaElements in FVal.Assigneds then
    begin
      AWriter.BeginTag('c:val');
      FVal.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:val');
{$ifdef XLS_WRITE_EXTLST}
  if (FExtLst <> Nil) and FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then
    begin
      AWriter.BeginTag('c:extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:extLst');
{$endif}
end;

constructor TCT_AreaSer.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 9;
  FAttributeCount := 0;
  FEG_SerShared := TEG_SerShared.Create(FOwner);
end;

destructor TCT_AreaSer.Destroy;
begin
  FEG_SerShared.Free;
  if FPictureOptions <> Nil then 
    FPictureOptions.Free;
  if FDPtXpgList <> Nil then 
    FDPtXpgList.Free;
  if FDLbls <> Nil then 
    FDLbls.Free;
  if FTrendlineXpgList <> Nil then 
    FTrendlineXpgList.Free;
  if FErrBarsXpgList <> Nil then 
    FErrBarsXpgList.Free;
  if FCat <> Nil then 
    FCat.Free;
  if FVal <> Nil then 
    FVal.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_AreaSer.Clear;
begin
  FAssigneds := [];
  FEG_SerShared.Clear;
  if FPictureOptions <> Nil then 
    FreeAndNil(FPictureOptions);
  if FDPtXpgList <> Nil then 
    FreeAndNil(FDPtXpgList);
  if FDLbls <> Nil then 
    FreeAndNil(FDLbls);
  if FTrendlineXpgList <> Nil then 
    FreeAndNil(FTrendlineXpgList);
  if FErrBarsXpgList <> Nil then 
    FreeAndNil(FErrBarsXpgList);
  if FCat <> Nil then 
    FreeAndNil(FCat);
  if FVal <> Nil then 
    FreeAndNil(FVal);
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
end;

procedure TCT_AreaSer.Assign(AItem: TCT_AreaSer);
begin
end;

procedure TCT_AreaSer.CopyTo(AItem: TCT_AreaSer);
begin
end;

function  TCT_AreaSer.Create_PictureOptions: TCT_PictureOptions;
begin
  Result := TCT_PictureOptions.Create(FOwner);
  FPictureOptions := Result;
end;

function  TCT_AreaSer.Create_DPtXpgList: TCT_DPtXpgList;
begin
  Result := TCT_DPtXpgList.Create(FOwner);
  FDPtXpgList := Result;
end;

function  TCT_AreaSer.Create_DLbls: TCT_DLbls;
begin
  Result := TCT_DLbls.Create(FOwner);
  FDLbls := Result;
end;

function  TCT_AreaSer.Create_TrendlineXpgList: TCT_TrendlineXpgList;
begin
  Result := TCT_TrendlineXpgList.Create(FOwner);
  FTrendlineXpgList := Result;
end;

function  TCT_AreaSer.Create_ErrBarsXpgList: TCT_ErrBarsXpgList;
begin
  Result := TCT_ErrBarsXpgList.Create(FOwner);
  FErrBarsXpgList := Result;
end;

function  TCT_AreaSer.Create_Cat: TCT_AxDataSource;
begin
  Result := TCT_AxDataSource.Create(FOwner);
  FCat := Result;
end;

function  TCT_AreaSer.Create_Val: TCT_NumDataSource;
begin
  Result := TCT_NumDataSource.Create(FOwner);
  FVal := Result;
end;

function  TCT_AreaSer.Create_ExtLst: TCT_ExtensionList;
begin
  Result := TCT_ExtensionList.Create(FOwner);
  FExtLst := Result;
end;

{ TCT_AreaSerXpgList }

function  TCT_AreaSerXpgList.GetItems(Index: integer): TCT_AreaSer;
begin
  Result := TCT_AreaSer(inherited Items[Index]);
end;

function  TCT_AreaSerXpgList.Add: TCT_AreaSer;
begin
  Result := TCT_AreaSer.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_AreaSerXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_AreaSerXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
    if xaElements in Items[i].Assigneds then
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
end;

procedure TCT_AreaSerXpgList.Assign(AItem: TCT_AreaSerXpgList);
begin
end;

procedure TCT_AreaSerXpgList.CopyTo(AItem: TCT_AreaSerXpgList);
begin
end;

{ TCT_LineSer }

function  TCT_LineSer.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FEG_SerShared.CheckAssigned);
  if FMarker <> Nil then 
    Inc(ElemsAssigned,FMarker.CheckAssigned);
  if FDPtXpgList <> Nil then 
    Inc(ElemsAssigned,FDPtXpgList.CheckAssigned);
  if FDLbls <> Nil then 
    Inc(ElemsAssigned,FDLbls.CheckAssigned);
  if FTrendlineXpgList <> Nil then 
    Inc(ElemsAssigned,FTrendlineXpgList.CheckAssigned);
  if FErrBars <> Nil then 
    Inc(ElemsAssigned,FErrBars.CheckAssigned);
  if FCat <> Nil then 
    Inc(ElemsAssigned,FCat.CheckAssigned);
  if FVal <> Nil then 
    Inc(ElemsAssigned,FVal.CheckAssigned);
  if FSmooth <> Nil then 
    Inc(ElemsAssigned,FSmooth.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_LineSer.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $0000031F: begin
      if FMarker = Nil then 
        FMarker := TCT_Marker.Create(FOwner);
      Result := FMarker;
    end;
    $000001C5: begin
      if FDPtXpgList = Nil then 
        FDPtXpgList := TCT_DPtXpgList.Create(FOwner);
      Result := FDPtXpgList.Add;
    end;
    $0000028E: begin
      if FDLbls = Nil then 
        FDLbls := TCT_DLbls.Create(FOwner);
      Result := FDLbls;
    end;
    $00000462: begin
      if FTrendlineXpgList = Nil then 
        FTrendlineXpgList := TCT_TrendlineXpgList.Create(FOwner);
      Result := FTrendlineXpgList.Add;
    end;
    $0000036E: begin
      if FErrBars = Nil then 
        FErrBars := TCT_ErrBars.Create(FOwner);
      Result := FErrBars;
    end;
    $000001D5: begin
      if FCat = Nil then 
        FCat := TCT_AxDataSource.Create(FOwner);
      Result := FCat;
    end;
    $000001E0: begin
      if FVal = Nil then 
        FVal := TCT_NumDataSource.Create(FOwner);
      Result := FVal;
    end;
    $00000337: begin
      if FSmooth = Nil then 
        FSmooth := TCT_Boolean.Create(FOwner);
      Result := FSmooth;
    end;
    $00000321: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
    begin
      Result := FEG_SerShared.HandleElement(AReader);
      if Result = Nil then 
        FOwner.Errors.Error(xemUnknownElement,AReader.QName);
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_LineSer.Write(AWriter: TXpgWriteXML);
begin
  FEG_SerShared.Write(AWriter);
  if (FMarker <> Nil) and FMarker.Assigned then 
    if xaElements in FMarker.Assigneds then
    begin
      AWriter.BeginTag('c:marker');
      FMarker.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:marker');
  if FDPtXpgList <> Nil then 
    FDPtXpgList.Write(AWriter,'c:dPt');
  if (FDLbls <> Nil) and FDLbls.Assigned then 
    if xaElements in FDLbls.Assigneds then
    begin
      AWriter.BeginTag('c:dLbls');
      FDLbls.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:dLbls');
  if FTrendlineXpgList <> Nil then 
    FTrendlineXpgList.Write(AWriter,'c:trendline');
  if (FErrBars <> Nil) and FErrBars.Assigned then 
    if xaElements in FErrBars.Assigneds then
    begin
      AWriter.BeginTag('c:errBars');
      FErrBars.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:errBars');
  if (FCat <> Nil) and FCat.Assigned then 
    if xaElements in FCat.Assigneds then
    begin
      AWriter.BeginTag('c:cat');
      FCat.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:cat');
  if (FVal <> Nil) and FVal.Assigned then 
    if xaElements in FVal.Assigneds then
    begin
      AWriter.BeginTag('c:val');
      FVal.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:val');
  if (FSmooth <> Nil) and FSmooth.Assigned then 
  begin
    FSmooth.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:smooth');
  end;
{$ifdef XLS_WRITE_EXTLST}
  if (FExtLst <> Nil) and FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then
    begin
      AWriter.BeginTag('c:extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:extLst');
{$endif}
end;

constructor TCT_LineSer.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 10;
  FAttributeCount := 0;
  FEG_SerShared := TEG_SerShared.Create(FOwner);
end;

destructor TCT_LineSer.Destroy;
begin
  FEG_SerShared.Free;
  if FMarker <> Nil then 
    FMarker.Free;
  if FDPtXpgList <> Nil then 
    FDPtXpgList.Free;
  if FDLbls <> Nil then 
    FDLbls.Free;
  if FTrendlineXpgList <> Nil then 
    FTrendlineXpgList.Free;
  if FErrBars <> Nil then 
    FErrBars.Free;
  if FCat <> Nil then 
    FCat.Free;
  if FVal <> Nil then 
    FVal.Free;
  if FSmooth <> Nil then 
    FSmooth.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_LineSer.Clear;
begin
  FAssigneds := [];
  FEG_SerShared.Clear;
  if FMarker <> Nil then 
    FreeAndNil(FMarker);
  if FDPtXpgList <> Nil then 
    FreeAndNil(FDPtXpgList);
  if FDLbls <> Nil then 
    FreeAndNil(FDLbls);
  if FTrendlineXpgList <> Nil then 
    FreeAndNil(FTrendlineXpgList);
  if FErrBars <> Nil then 
    FreeAndNil(FErrBars);
  if FCat <> Nil then 
    FreeAndNil(FCat);
  if FVal <> Nil then 
    FreeAndNil(FVal);
  if FSmooth <> Nil then 
    FreeAndNil(FSmooth);
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
end;

procedure TCT_LineSer.Assign(AItem: TCT_LineSer);
begin
end;

procedure TCT_LineSer.CopyTo(AItem: TCT_LineSer);
begin
end;

function  TCT_LineSer.Create_Marker: TCT_Marker;
begin
  Result := TCT_Marker.Create(FOwner);
  FMarker := Result;
end;

function  TCT_LineSer.Create_DPtXpgList: TCT_DPtXpgList;
begin
  Result := TCT_DPtXpgList.Create(FOwner);
  FDPtXpgList := Result;
end;

function  TCT_LineSer.Create_DLbls: TCT_DLbls;
begin
  Result := TCT_DLbls.Create(FOwner);
  FDLbls := Result;
end;

function  TCT_LineSer.Create_TrendlineXpgList: TCT_TrendlineXpgList;
begin
  Result := TCT_TrendlineXpgList.Create(FOwner);
  FTrendlineXpgList := Result;
end;

function  TCT_LineSer.Create_ErrBars: TCT_ErrBars;
begin
  Result := TCT_ErrBars.Create(FOwner);
  FErrBars := Result;
end;

function  TCT_LineSer.Create_Cat: TCT_AxDataSource;
begin
  Result := TCT_AxDataSource.Create(FOwner);
  FCat := Result;
end;

function  TCT_LineSer.Create_Val: TCT_NumDataSource;
begin
  Result := TCT_NumDataSource.Create(FOwner);
  FVal := Result;
end;

function  TCT_LineSer.Create_Smooth: TCT_Boolean;
begin
  Result := TCT_Boolean.Create(FOwner);
  FSmooth := Result;
end;

function  TCT_LineSer.Create_ExtLst: TCT_ExtensionList;
begin
  Result := TCT_ExtensionList.Create(FOwner);
  FExtLst := Result;
end;

{ TCT_LineSerXpgList }

function  TCT_LineSerXpgList.GetItems(Index: integer): TCT_LineSer;
begin
  Result := TCT_LineSer(inherited Items[Index]);
end;

function  TCT_LineSerXpgList.Add: TCT_LineSer;
begin
  Result := TCT_LineSer.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_LineSerXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_LineSerXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
    if xaElements in Items[i].Assigneds then
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
end;

procedure TCT_LineSerXpgList.Assign(AItem: TCT_LineSerXpgList);
begin
end;

procedure TCT_LineSerXpgList.CopyTo(AItem: TCT_LineSerXpgList);
begin
end;

{ TCT_GapAmount }

function  TCT_GapAmount.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FVal <> 150 then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_GapAmount.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_GapAmount.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FVal <> 150 then 
    AWriter.AddAttribute('val',XmlIntToStr(FVal));
end;

procedure TCT_GapAmount.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then 
    FVal := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_GapAmount.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := 150;
end;

destructor TCT_GapAmount.Destroy;
begin
end;

procedure TCT_GapAmount.Clear;
begin
  FAssigneds := [];
  FVal := 150;
end;

procedure TCT_GapAmount.Assign(AItem: TCT_GapAmount);
begin
end;

procedure TCT_GapAmount.CopyTo(AItem: TCT_GapAmount);
begin
end;

{ TCT_UpDownBar }

function  TCT_UpDownBar.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FSpPr <> Nil then 
    Inc(ElemsAssigned,FSpPr.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_UpDownBar.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'c:spPr' then 
  begin
    if FSpPr = Nil then 
      FSpPr := TCT_ShapeProperties.Create(FOwner);
    Result := FSpPr;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_UpDownBar.Write(AWriter: TXpgWriteXML);
begin
  if (FSpPr <> Nil) and FSpPr.Assigned then 
  begin
    FSpPr.WriteAttributes(AWriter);
    if xaElements in FSpPr.Assigneds then
    begin
      AWriter.BeginTag('c:spPr');
      FSpPr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:spPr');
  end;
end;

constructor TCT_UpDownBar.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
end;

destructor TCT_UpDownBar.Destroy;
begin
  if FSpPr <> Nil then 
    FSpPr.Free;
end;

procedure TCT_UpDownBar.Clear;
begin
  FAssigneds := [];
  if FSpPr <> Nil then 
    FreeAndNil(FSpPr);
end;

procedure TCT_UpDownBar.Assign(AItem: TCT_UpDownBar);
begin
end;

procedure TCT_UpDownBar.CopyTo(AItem: TCT_UpDownBar);
begin
end;

function  TCT_UpDownBar.Create_SpPr: TCT_ShapeProperties;
begin
  Result := TCT_ShapeProperties.Create(FOwner);
  FSpPr := Result;
end;

{ TCT_PieSer }

function  TCT_PieSer.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FEG_SerShared.CheckAssigned);
  if FExplosion <> Nil then 
    Inc(ElemsAssigned,FExplosion.CheckAssigned);
  if FDPtXpgList <> Nil then 
    Inc(ElemsAssigned,FDPtXpgList.CheckAssigned);
  if FDLbls <> Nil then 
    Inc(ElemsAssigned,FDLbls.CheckAssigned);
  if FCat <> Nil then 
    Inc(ElemsAssigned,FCat.CheckAssigned);
  if FVal <> Nil then 
    Inc(ElemsAssigned,FVal.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_PieSer.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $0000047E: begin
      if FExplosion = Nil then 
        FExplosion := TCT_UnsignedInt.Create(FOwner);
      Result := FExplosion;
    end;
    $000001C5: begin
      if FDPtXpgList = Nil then 
        FDPtXpgList := TCT_DPtXpgList.Create(FOwner);
      Result := FDPtXpgList.Add;
    end;
    $0000028E: begin
      if FDLbls = Nil then 
        FDLbls := TCT_DLbls.Create(FOwner);
      Result := FDLbls;
    end;
    $000001D5: begin
      if FCat = Nil then 
        FCat := TCT_AxDataSource.Create(FOwner);
      Result := FCat;
    end;
    $000001E0: begin
      if FVal = Nil then 
        FVal := TCT_NumDataSource.Create(FOwner);
      Result := FVal;
    end;
    $00000321: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
    begin
      Result := FEG_SerShared.HandleElement(AReader);
      if Result = Nil then 
        FOwner.Errors.Error(xemUnknownElement,AReader.QName);
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_PieSer.Write(AWriter: TXpgWriteXML);
begin
  FEG_SerShared.Write(AWriter);
  if (FExplosion <> Nil) and FExplosion.Assigned then 
  begin
    FExplosion.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:explosion');
  end;
  if FDPtXpgList <> Nil then 
    FDPtXpgList.Write(AWriter,'c:dPt');
  if (FDLbls <> Nil) and FDLbls.Assigned then 
    if xaElements in FDLbls.Assigneds then
    begin
      AWriter.BeginTag('c:dLbls');
      FDLbls.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:dLbls');
  if (FCat <> Nil) and FCat.Assigned then 
    if xaElements in FCat.Assigneds then
    begin
      AWriter.BeginTag('c:cat');
      FCat.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:cat');
  if (FVal <> Nil) and FVal.Assigned then 
    if xaElements in FVal.Assigneds then
    begin
      AWriter.BeginTag('c:val');
      FVal.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:val');
{$ifdef XLS_WRITE_EXTLST}
  if (FExtLst <> Nil) and FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then
    begin
      AWriter.BeginTag('c:extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:extLst');
{$endif}
end;

constructor TCT_PieSer.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 7;
  FAttributeCount := 0;
  FEG_SerShared := TEG_SerShared.Create(FOwner);
end;

destructor TCT_PieSer.Destroy;
begin
  FEG_SerShared.Free;
  if FExplosion <> Nil then 
    FExplosion.Free;
  if FDPtXpgList <> Nil then 
    FDPtXpgList.Free;
  if FDLbls <> Nil then 
    FDLbls.Free;
  if FCat <> Nil then 
    FCat.Free;
  if FVal <> Nil then 
    FVal.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_PieSer.Clear;
begin
  FAssigneds := [];
  FEG_SerShared.Clear;
  if FExplosion <> Nil then 
    FreeAndNil(FExplosion);
  if FDPtXpgList <> Nil then 
    FreeAndNil(FDPtXpgList);
  if FDLbls <> Nil then 
    FreeAndNil(FDLbls);
  if FCat <> Nil then 
    FreeAndNil(FCat);
  if FVal <> Nil then 
    FreeAndNil(FVal);
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
end;

procedure TCT_PieSer.Assign(AItem: TCT_PieSer);
begin
end;

procedure TCT_PieSer.CopyTo(AItem: TCT_PieSer);
begin
end;

function  TCT_PieSer.Create_Explosion: TCT_UnsignedInt;
begin
  Result := TCT_UnsignedInt.Create(FOwner);
  FExplosion := Result;
end;

function  TCT_PieSer.Create_DPtXpgList: TCT_DPtXpgList;
begin
  Result := TCT_DPtXpgList.Create(FOwner);
  FDPtXpgList := Result;
end;

function  TCT_PieSer.Create_DLbls: TCT_DLbls;
begin
  Result := TCT_DLbls.Create(FOwner);
  FDLbls := Result;
end;

function  TCT_PieSer.Create_Cat: TCT_AxDataSource;
begin
  Result := TCT_AxDataSource.Create(FOwner);
  FCat := Result;
end;

function  TCT_PieSer.Create_Val: TCT_NumDataSource;
begin
  Result := TCT_NumDataSource.Create(FOwner);
  FVal := Result;
end;

function  TCT_PieSer.Create_ExtLst: TCT_ExtensionList;
begin
  Result := TCT_ExtensionList.Create(FOwner);
  FExtLst := Result;
end;

{ TCT_PieSerXpgList }

function  TCT_PieSerXpgList.GetItems(Index: integer): TCT_PieSer;
begin
  Result := TCT_PieSer(inherited Items[Index]);
end;

function  TCT_PieSerXpgList.Add: TCT_PieSer;
begin
  Result := TCT_PieSer.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_PieSerXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_PieSerXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
    if xaElements in Items[i].Assigneds then
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag(AName);
end;

procedure TCT_PieSerXpgList.Assign(AItem: TCT_PieSerXpgList);
begin
end;

procedure TCT_PieSerXpgList.CopyTo(AItem: TCT_PieSerXpgList);
begin
end;

{ TCT_BarDir }

function  TCT_BarDir.CheckAssigned: integer;
begin
  Result := 1;
  FAssigneds := [xaAttributes];
end;

procedure TCT_BarDir.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_BarDir.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('val',StrTST_BarDir[Integer(FVal)]);
end;

procedure TCT_BarDir.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then
    FVal := TST_BarDir(StrToEnum('stbd' + AAttributes.Values[0]))
  else
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_BarDir.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := stbdCol;
end;

destructor TCT_BarDir.Destroy;
begin
end;

procedure TCT_BarDir.Clear;
begin
  FAssigneds := [];
  FVal := stbdCol;
end;

procedure TCT_BarDir.Assign(AItem: TCT_BarDir);
begin
end;

procedure TCT_BarDir.CopyTo(AItem: TCT_BarDir);
begin
end;

{ TCT_BarGrouping }

function  TCT_BarGrouping.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 1;
  FAssigneds := [];
//  if FVal <> stbgClustered then
//    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_BarGrouping.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_BarGrouping.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('val',StrTST_BarGrouping[Integer(FVal)]);
end;

procedure TCT_BarGrouping.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then 
    FVal := TST_BarGrouping(StrToEnum('stbg' + AAttributes.Values[0]))
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_BarGrouping.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := stbgClustered;
end;

destructor TCT_BarGrouping.Destroy;
begin
end;

procedure TCT_BarGrouping.Clear;
begin
  FAssigneds := [];
  FVal := stbgClustered;
end;

procedure TCT_BarGrouping.Assign(AItem: TCT_BarGrouping);
begin
end;

procedure TCT_BarGrouping.CopyTo(AItem: TCT_BarGrouping);
begin
end;

{ TCT_BarSer }

function  TCT_BarSer.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FEG_SerShared.CheckAssigned);
  if FInvertIfNegative <> Nil then 
    Inc(ElemsAssigned,FInvertIfNegative.CheckAssigned);
  if FPictureOptions <> Nil then 
    Inc(ElemsAssigned,FPictureOptions.CheckAssigned);
  if FDPtXpgList <> Nil then 
    Inc(ElemsAssigned,FDPtXpgList.CheckAssigned);
  if FDLbls <> Nil then 
    Inc(ElemsAssigned,FDLbls.CheckAssigned);
  if FTrendlineXpgList <> Nil then 
    Inc(ElemsAssigned,FTrendlineXpgList.CheckAssigned);
  if FErrBars <> Nil then 
    Inc(ElemsAssigned,FErrBars.CheckAssigned);
  if FCat <> Nil then 
    Inc(ElemsAssigned,FCat.CheckAssigned);
  if FVal <> Nil then 
    Inc(ElemsAssigned,FVal.CheckAssigned);
  if FShape <> Nil then 
    Inc(ElemsAssigned,FShape.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_BarSer.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $00000717: begin
      if FInvertIfNegative = Nil then 
        FInvertIfNegative := TCT_Boolean.Create(FOwner);
      Result := FInvertIfNegative;
    end;
    $00000685: begin
      if FPictureOptions = Nil then 
        FPictureOptions := TCT_PictureOptions.Create(FOwner);
      Result := FPictureOptions;
    end;
    $000001C5: begin
      if FDPtXpgList = Nil then 
        FDPtXpgList := TCT_DPtXpgList.Create(FOwner);
      Result := FDPtXpgList.Add;
    end;
    $0000028E: begin
      if FDLbls = Nil then 
        FDLbls := TCT_DLbls.Create(FOwner);
      Result := FDLbls;
    end;
    $00000462: begin
      if FTrendlineXpgList = Nil then 
        FTrendlineXpgList := TCT_TrendlineXpgList.Create(FOwner);
      Result := FTrendlineXpgList.Add;
    end;
    $0000036E: begin
      if FErrBars = Nil then 
        FErrBars := TCT_ErrBars.Create(FOwner);
      Result := FErrBars;
    end;
    $000001D5: begin
      if FCat = Nil then 
        FCat := TCT_AxDataSource.Create(FOwner);
      Result := FCat;
    end;
    $000001E0: begin
      if FVal = Nil then 
        FVal := TCT_NumDataSource.Create(FOwner);
      Result := FVal;
    end;
    $000002AE: begin
      if FShape = Nil then 
        FShape := TCT_Shape.Create(FOwner);
      Result := FShape;
    end;
    $00000321: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
    begin
      Result := FEG_SerShared.HandleElement(AReader);
      if Result = Nil then 
        FOwner.Errors.Error(xemUnknownElement,AReader.QName);
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_BarSer.Write(AWriter: TXpgWriteXML);
begin
  FEG_SerShared.Write(AWriter);
  if (FInvertIfNegative <> Nil) and FInvertIfNegative.Assigned then
  begin
    FInvertIfNegative.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:invertIfNegative');
  end;
  if (FPictureOptions <> Nil) and FPictureOptions.Assigned then
    if xaElements in FPictureOptions.Assigneds then
    begin
      AWriter.BeginTag('c:pictureOptions');
      FPictureOptions.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:pictureOptions');
  if FDPtXpgList <> Nil then
    FDPtXpgList.Write(AWriter,'c:dPt');
  if (FDLbls <> Nil) and FDLbls.Assigned then 
    if xaElements in FDLbls.Assigneds then
    begin
      AWriter.BeginTag('c:dLbls');
      FDLbls.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:dLbls');
  if FTrendlineXpgList <> Nil then 
    FTrendlineXpgList.Write(AWriter,'c:trendline');
  if (FErrBars <> Nil) and FErrBars.Assigned then 
    if xaElements in FErrBars.Assigneds then
    begin
      AWriter.BeginTag('c:errBars');
      FErrBars.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:errBars');
  if (FCat <> Nil) and FCat.Assigned then 
    if xaElements in FCat.Assigneds then
    begin
      AWriter.BeginTag('c:cat');
      FCat.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:cat');
  if (FVal <> Nil) and FVal.Assigned then 
    if xaElements in FVal.Assigneds then
    begin
      AWriter.BeginTag('c:val');
      FVal.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:val');
  if (FShape <> Nil) and FShape.Assigned then 
  begin
    FShape.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:shape');
  end;
{$ifdef XLS_WRITE_EXTLST}
  if (FExtLst <> Nil) and FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then
    begin
      AWriter.BeginTag('c:extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:extLst');
{$endif}
end;

constructor TCT_BarSer.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 11;
  FAttributeCount := 0;
  FEG_SerShared := TEG_SerShared.Create(FOwner);
end;

destructor TCT_BarSer.Destroy;
begin
  FEG_SerShared.Free;
  if FInvertIfNegative <> Nil then 
    FInvertIfNegative.Free;
  if FPictureOptions <> Nil then 
    FPictureOptions.Free;
  if FDPtXpgList <> Nil then 
    FDPtXpgList.Free;
  if FDLbls <> Nil then 
    FDLbls.Free;
  if FTrendlineXpgList <> Nil then 
    FTrendlineXpgList.Free;
  if FErrBars <> Nil then 
    FErrBars.Free;
  if FCat <> Nil then 
    FCat.Free;
  if FVal <> Nil then 
    FVal.Free;
  if FShape <> Nil then 
    FShape.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_BarSer.Clear;
begin
  FAssigneds := [];
  FEG_SerShared.Clear;
  if FInvertIfNegative <> Nil then 
    FreeAndNil(FInvertIfNegative);
  if FPictureOptions <> Nil then 
    FreeAndNil(FPictureOptions);
  if FDPtXpgList <> Nil then 
    FreeAndNil(FDPtXpgList);
  if FDLbls <> Nil then 
    FreeAndNil(FDLbls);
  if FTrendlineXpgList <> Nil then 
    FreeAndNil(FTrendlineXpgList);
  if FErrBars <> Nil then 
    FreeAndNil(FErrBars);
  if FCat <> Nil then 
    FreeAndNil(FCat);
  if FVal <> Nil then 
    FreeAndNil(FVal);
  if FShape <> Nil then 
    FreeAndNil(FShape);
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
end;

procedure TCT_BarSer.Assign(AItem: TCT_BarSer);
begin
end;

procedure TCT_BarSer.CopyTo(AItem: TCT_BarSer);
begin
end;

function  TCT_BarSer.Create_InvertIfNegative: TCT_Boolean;
begin
  Result := TCT_Boolean.Create(FOwner);
  FInvertIfNegative := Result;
end;

function  TCT_BarSer.Create_PictureOptions: TCT_PictureOptions;
begin
  Result := TCT_PictureOptions.Create(FOwner);
  FPictureOptions := Result;
end;

function  TCT_BarSer.Create_DPtXpgList: TCT_DPtXpgList;
begin
  Result := TCT_DPtXpgList.Create(FOwner);
  FDPtXpgList := Result;
end;

function  TCT_BarSer.Create_DLbls: TCT_DLbls;
begin
  Result := TCT_DLbls.Create(FOwner);
  FDLbls := Result;
end;

function  TCT_BarSer.Create_TrendlineXpgList: TCT_TrendlineXpgList;
begin
  Result := TCT_TrendlineXpgList.Create(FOwner);
  FTrendlineXpgList := Result;
end;

function  TCT_BarSer.Create_ErrBars: TCT_ErrBars;
begin
  Result := TCT_ErrBars.Create(FOwner);
  FErrBars := Result;
end;

function  TCT_BarSer.Create_Cat: TCT_AxDataSource;
begin
  Result := TCT_AxDataSource.Create(FOwner);
  FCat := Result;
end;

function  TCT_BarSer.Create_Val: TCT_NumDataSource;
begin
  Result := TCT_NumDataSource.Create(FOwner);
  FVal := Result;
end;

function  TCT_BarSer.Create_Shape: TCT_Shape;
begin
  Result := TCT_Shape.Create(FOwner);
  FShape := Result;
end;

function  TCT_BarSer.Create_ExtLst: TCT_ExtensionList;
begin
  Result := TCT_ExtensionList.Create(FOwner);
  FExtLst := Result;
end;

{ TCT_BarSerXpgList }

function  TCT_BarSerXpgList.GetItems(Index: integer): TCT_BarSer;
begin
  Result := TCT_BarSer(inherited Items[Index]);
end;

function  TCT_BarSerXpgList.Add: TCT_BarSer;
begin
  Result := TCT_BarSer.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_BarSerXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_BarSerXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
    if xaElements in Items[i].Assigneds then
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
end;

procedure TCT_BarSerXpgList.Assign(AItem: TCT_BarSerXpgList);
begin
end;

procedure TCT_BarSerXpgList.CopyTo(AItem: TCT_BarSerXpgList);
begin
end;

{ TCT_SurfaceSer }

function  TCT_SurfaceSer.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FEG_SerShared.CheckAssigned);
  if FCat <> Nil then 
    Inc(ElemsAssigned,FCat.CheckAssigned);
  if FVal <> Nil then 
    Inc(ElemsAssigned,FVal.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_SurfaceSer.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $000001D5: begin
      if FCat = Nil then 
        FCat := TCT_AxDataSource.Create(FOwner);
      Result := FCat;
    end;
    $000001E0: begin
      if FVal = Nil then 
        FVal := TCT_NumDataSource.Create(FOwner);
      Result := FVal;
    end;
    $00000321: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
    begin
      Result := FEG_SerShared.HandleElement(AReader);
      if Result = Nil then 
        FOwner.Errors.Error(xemUnknownElement,AReader.QName);
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_SurfaceSer.Write(AWriter: TXpgWriteXML);
begin
  FEG_SerShared.Write(AWriter);
  if (FCat <> Nil) and FCat.Assigned then 
    if xaElements in FCat.Assigneds then
    begin
      AWriter.BeginTag('c:cat');
      FCat.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:cat');
  if (FVal <> Nil) and FVal.Assigned then 
    if xaElements in FVal.Assigneds then
    begin
      AWriter.BeginTag('c:val');
      FVal.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:val');
{$ifdef XLS_WRITE_EXTLST}
  if (FExtLst <> Nil) and FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then
    begin
      AWriter.BeginTag('c:extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:extLst');
{$endif}
end;

constructor TCT_SurfaceSer.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 4;
  FAttributeCount := 0;
  FEG_SerShared := TEG_SerShared.Create(FOwner);
end;

destructor TCT_SurfaceSer.Destroy;
begin
  FEG_SerShared.Free;
  if FCat <> Nil then 
    FCat.Free;
  if FVal <> Nil then 
    FVal.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_SurfaceSer.Clear;
begin
  FAssigneds := [];
  FEG_SerShared.Clear;
  if FCat <> Nil then 
    FreeAndNil(FCat);
  if FVal <> Nil then 
    FreeAndNil(FVal);
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
end;

procedure TCT_SurfaceSer.Assign(AItem: TCT_SurfaceSer);
begin
end;

procedure TCT_SurfaceSer.CopyTo(AItem: TCT_SurfaceSer);
begin
end;

function  TCT_SurfaceSer.Create_Cat: TCT_AxDataSource;
begin
  Result := TCT_AxDataSource.Create(FOwner);
  FCat := Result;
end;

function  TCT_SurfaceSer.Create_Val: TCT_NumDataSource;
begin
  Result := TCT_NumDataSource.Create(FOwner);
  FVal := Result;
end;

function  TCT_SurfaceSer.Create_ExtLst: TCT_ExtensionList;
begin
  Result := TCT_ExtensionList.Create(FOwner);
  FExtLst := Result;
end;

{ TCT_SurfaceSerXpgList }

function  TCT_SurfaceSerXpgList.GetItems(Index: integer): TCT_SurfaceSer;
begin
  Result := TCT_SurfaceSer(inherited Items[Index]);
end;

function  TCT_SurfaceSerXpgList.Add: TCT_SurfaceSer;
begin
  Result := TCT_SurfaceSer.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_SurfaceSerXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_SurfaceSerXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
    if xaElements in Items[i].Assigneds then
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
end;

procedure TCT_SurfaceSerXpgList.Assign(AItem: TCT_SurfaceSerXpgList);
begin
end;

procedure TCT_SurfaceSerXpgList.CopyTo(AItem: TCT_SurfaceSerXpgList);
begin
end;

{ TCT_BandFmts }

function  TCT_BandFmts.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FBandFmtXpgList <> Nil then 
    Inc(ElemsAssigned,FBandFmtXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_BandFmts.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'c:bandFmt' then 
  begin
    if FBandFmtXpgList = Nil then 
      FBandFmtXpgList := TCT_BandFmtXpgList.Create(FOwner);
    Result := FBandFmtXpgList.Add;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_BandFmts.Write(AWriter: TXpgWriteXML);
begin
  if FBandFmtXpgList <> Nil then 
    FBandFmtXpgList.Write(AWriter,'c:bandFmt');
end;

constructor TCT_BandFmts.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
end;

destructor TCT_BandFmts.Destroy;
begin
  if FBandFmtXpgList <> Nil then 
    FBandFmtXpgList.Free;
end;

procedure TCT_BandFmts.Clear;
begin
  FAssigneds := [];
  if FBandFmtXpgList <> Nil then 
    FreeAndNil(FBandFmtXpgList);
end;

procedure TCT_BandFmts.Assign(AItem: TCT_BandFmts);
begin
end;

procedure TCT_BandFmts.CopyTo(AItem: TCT_BandFmts);
begin
end;

function  TCT_BandFmts.Create_BandFmtXpgList: TCT_BandFmtXpgList;
begin
  Result := TCT_BandFmtXpgList.Create(FOwner);
  FBandFmtXpgList := Result;
end;

{ TCT_Scaling }

function  TCT_Scaling.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FLogBase <> Nil then 
    Inc(ElemsAssigned,FLogBase.CheckAssigned);
  if FOrientation <> Nil then 
    Inc(ElemsAssigned,FOrientation.CheckAssigned);
  if FMax <> Nil then 
    Inc(ElemsAssigned,FMax.CheckAssigned);
  if FMin <> Nil then 
    Inc(ElemsAssigned,FMin.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_Scaling.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $0000035A: begin
      if FLogBase = Nil then 
        FLogBase := TCT_LogBase.Create(FOwner);
      Result := FLogBase;
    end;
    $00000549: begin
      if FOrientation = Nil then 
        FOrientation := TCT_Orientation.Create(FOwner);
      Result := FOrientation;
    end;
    $000001E3: begin
      if FMax = Nil then 
        FMax := TCT_Double.Create(FOwner);
      Result := FMax;
    end;
    $000001E1: begin
      if FMin = Nil then 
        FMin := TCT_Double.Create(FOwner);
      Result := FMin;
    end;
    $00000321: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_Scaling.Write(AWriter: TXpgWriteXML);
begin
  if (FLogBase <> Nil) and FLogBase.Assigned then 
  begin
    FLogBase.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:logBase');
  end;
//  if (FOrientation <> Nil) and FOrientation.Assigned then
//  begin
    FOrientation.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:orientation');
//  end;
  if (FMax <> Nil) and FMax.Assigned then 
  begin
    FMax.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:max');
  end;
  if (FMin <> Nil) and FMin.Assigned then 
  begin
    FMin.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:min');
  end;
{$ifdef XLS_WRITE_EXTLST}
  if (FExtLst <> Nil) and FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then
    begin
      AWriter.BeginTag('c:extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:extLst');
{$endif}
end;

constructor TCT_Scaling.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 5;
  FAttributeCount := 0;
end;

destructor TCT_Scaling.Destroy;
begin
  if FLogBase <> Nil then 
    FLogBase.Free;
  if FOrientation <> Nil then 
    FOrientation.Free;
  if FMax <> Nil then 
    FMax.Free;
  if FMin <> Nil then 
    FMin.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_Scaling.Clear;
begin
  FAssigneds := [];
  if FLogBase <> Nil then 
    FreeAndNil(FLogBase);
  if FOrientation <> Nil then 
    FreeAndNil(FOrientation);
  if FMax <> Nil then 
    FreeAndNil(FMax);
  if FMin <> Nil then 
    FreeAndNil(FMin);
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
end;

procedure TCT_Scaling.Assign(AItem: TCT_Scaling);
begin
end;

procedure TCT_Scaling.CopyTo(AItem: TCT_Scaling);
begin
end;

function  TCT_Scaling.Create_LogBase: TCT_LogBase;
begin
  Result := TCT_LogBase.Create(FOwner);
  FLogBase := Result;
end;

function  TCT_Scaling.Create_Orientation: TCT_Orientation;
begin
  Result := TCT_Orientation.Create(FOwner);
  FOrientation := Result;
end;

function  TCT_Scaling.Create_Max: TCT_Double;
begin
  Result := TCT_Double.Create(FOwner);
  FMax := Result;
end;

function  TCT_Scaling.Create_Min: TCT_Double;
begin
  Result := TCT_Double.Create(FOwner);
  FMin := Result;
end;

function  TCT_Scaling.Create_ExtLst: TCT_ExtensionList;
begin
  Result := TCT_ExtensionList.Create(FOwner);
  FExtLst := Result;
end;

{ TCT_AxPos }

function  TCT_AxPos.CheckAssigned: integer;
begin
  FAssigneds := [xaAttributes];
  if FVal <> TST_AxPos(XPG_UNKNOWN_ENUM) then
    Result := 1
  else
    Result := 0;
end;

procedure TCT_AxPos.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_AxPos.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FVal <> TST_AxPos(XPG_UNKNOWN_ENUM) then
    AWriter.AddAttribute('val',StrTST_AxPos[Integer(FVal)]);
end;

procedure TCT_AxPos.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then
    FVal := TST_AxPos(StrToEnum('stap' + AAttributes.Values[0]))
  else
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_AxPos.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  Clear;
end;

destructor TCT_AxPos.Destroy;
begin
end;

procedure TCT_AxPos.Clear;
begin
  FAssigneds := [];
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := TST_AxPos(XPG_UNKNOWN_ENUM);
end;

procedure TCT_AxPos.Assign(AItem: TCT_AxPos);
begin
end;

procedure TCT_AxPos.CopyTo(AItem: TCT_AxPos);
begin
end;

{ TCT_Title }

function  TCT_Title.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FTx <> Nil then 
    Inc(ElemsAssigned,FTx.CheckAssigned);
  if FLayout <> Nil then 
    Inc(ElemsAssigned,FLayout.CheckAssigned);
  if FOverlay <> Nil then 
    Inc(ElemsAssigned,FOverlay.CheckAssigned);
  if FSpPr <> Nil then 
    Inc(ElemsAssigned,FSpPr.CheckAssigned);
  if FTxPr <> Nil then 
    Inc(ElemsAssigned,FTxPr.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_Title.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $00000189: begin
      if FTx = Nil then 
        FTx := TCT_Tx.Create(FOwner);
      Result := FTx;
    end;
    $0000033B: begin
      if FLayout = Nil then 
        FLayout := TCT_Layout.Create(FOwner);
      Result := FLayout;
    end;
    $0000039F: begin
      if FOverlay = Nil then 
        FOverlay := TCT_Boolean.Create(FOwner);
      Result := FOverlay;
    end;
    $00000242: begin
      if FSpPr = Nil then 
        FSpPr := TCT_ShapeProperties.Create(FOwner);
      Result := FSpPr;
    end;
    $0000024B: begin
      if FTxPr = Nil then 
        FTxPr := TCT_TextBody.Create(FOwner);
      Result := FTxPr;
    end;
    $00000321: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_Title.Write(AWriter: TXpgWriteXML);
begin
  if (FTx <> Nil) and FTx.Assigned then 
    if xaElements in FTx.Assigneds then
    begin
      AWriter.BeginTag('c:tx');
      FTx.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:tx');
  if (FLayout <> Nil) and FLayout.Assigned then 
    if xaElements in FLayout.Assigneds then
    begin
      AWriter.BeginTag('c:layout');
      FLayout.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:layout');
  if (FOverlay <> Nil) and FOverlay.Assigned then 
  begin
    FOverlay.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:overlay');
  end;
  if (FSpPr <> Nil) and FSpPr.Assigned then 
  begin
    FSpPr.WriteAttributes(AWriter);
    if xaElements in FSpPr.Assigneds then
    begin
      AWriter.BeginTag('c:spPr');
      FSpPr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:spPr');
  end;
  if (FTxPr <> Nil) and FTxPr.Assigned then 
    if xaElements in FTxPr.Assigneds then
    begin
      AWriter.BeginTag('c:txPr');
      FTxPr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:txPr');
{$ifdef XLS_WRITE_EXTLST}
  if (FExtLst <> Nil) and FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then
    begin
      AWriter.BeginTag('c:extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:extLst');
{$endif}
end;

constructor TCT_Title.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 6;
  FAttributeCount := 0;
end;

destructor TCT_Title.Destroy;
begin
  if FTx <> Nil then 
    FTx.Free;
  if FLayout <> Nil then 
    FLayout.Free;
  if FOverlay <> Nil then 
    FOverlay.Free;
  if FSpPr <> Nil then 
    FSpPr.Free;
  if FTxPr <> Nil then 
    FTxPr.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_Title.Clear;
begin
  FAssigneds := [];
  if FTx <> Nil then 
    FreeAndNil(FTx);
  if FLayout <> Nil then 
    FreeAndNil(FLayout);
  if FOverlay <> Nil then 
    FreeAndNil(FOverlay);
  if FSpPr <> Nil then 
    FreeAndNil(FSpPr);
  if FTxPr <> Nil then 
    FreeAndNil(FTxPr);
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
end;

procedure TCT_Title.Assign(AItem: TCT_Title);
begin
end;

procedure TCT_Title.CopyTo(AItem: TCT_Title);
begin
end;

function  TCT_Title.Create_Tx: TCT_Tx;
begin
  Result := TCT_Tx.Create(FOwner);
  FTx := Result;
end;

function  TCT_Title.Create_Layout: TCT_Layout;
begin
  Result := TCT_Layout.Create(FOwner);
  FLayout := Result;
end;

function  TCT_Title.Create_Overlay: TCT_Boolean;
begin
  Result := TCT_Boolean.Create(FOwner);
  FOverlay := Result;
end;

function  TCT_Title.Create_SpPr: TCT_ShapeProperties;
begin
  Result := TCT_ShapeProperties.Create(FOwner);
  FSpPr := Result;
end;

function  TCT_Title.Create_TxPr: TCT_TextBody;
begin
  Result := TCT_TextBody.Create(FOwner);
  FTxPr := Result;
end;

function  TCT_Title.Create_ExtLst: TCT_ExtensionList;
begin
  Result := TCT_ExtensionList.Create(FOwner);
  FExtLst := Result;
end;

{ TCT_TickMark }

function  TCT_TickMark.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FVal <> sttmCross then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_TickMark.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_TickMark.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FVal <> sttmCross then 
    AWriter.AddAttribute('val',StrTST_TickMark[Integer(FVal)]);
end;

procedure TCT_TickMark.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then 
    FVal := TST_TickMark(StrToEnum('sttm' + AAttributes.Values[0]))
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_TickMark.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := sttmCross;
end;

destructor TCT_TickMark.Destroy;
begin
end;

procedure TCT_TickMark.Clear;
begin
  FAssigneds := [];
  FVal := sttmCross;
end;

procedure TCT_TickMark.Assign(AItem: TCT_TickMark);
begin
end;

procedure TCT_TickMark.CopyTo(AItem: TCT_TickMark);
begin
end;

{ TCT_TickLblPos }

function  TCT_TickLblPos.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FVal <> sttlpNextTo then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_TickLblPos.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_TickLblPos.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FVal <> sttlpNextTo then 
    AWriter.AddAttribute('val',StrTST_TickLblPos[Integer(FVal)]);
end;

procedure TCT_TickLblPos.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then 
    FVal := TST_TickLblPos(StrToEnum('sttlp' + AAttributes.Values[0]))
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_TickLblPos.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := sttlpNextTo;
end;

destructor TCT_TickLblPos.Destroy;
begin
end;

procedure TCT_TickLblPos.Clear;
begin
  FAssigneds := [];
  FVal := sttlpNextTo;
end;

procedure TCT_TickLblPos.Assign(AItem: TCT_TickLblPos);
begin
end;

procedure TCT_TickLblPos.CopyTo(AItem: TCT_TickLblPos);
begin
end;

{ TCT_Crosses }

function  TCT_Crosses.CheckAssigned: integer;
begin
  FAssigneds := [xaAttributes];
  if FVal <> TST_Crosses(XPG_UNKNOWN_ENUM) then
    Result := 1
  else
    Result := 0;
end;

procedure TCT_Crosses.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_Crosses.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FVal <> TST_Crosses(XPG_UNKNOWN_ENUM) then
    AWriter.AddAttribute('val',StrTST_Crosses[Integer(FVal)]);
end;

procedure TCT_Crosses.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then
    FVal := TST_Crosses(StrToEnum('stc' + AAttributes.Values[0]))
  else
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_Crosses.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  Clear;
end;

destructor TCT_Crosses.Destroy;
begin
end;

procedure TCT_Crosses.Clear;
begin
  FAssigneds := [];
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := TST_Crosses(XPG_UNKNOWN_ENUM);
end;

procedure TCT_Crosses.Assign(AItem: TCT_Crosses);
begin
end;

procedure TCT_Crosses.CopyTo(AItem: TCT_Crosses);
begin
end;

{ TCT_BuiltInUnit }

function  TCT_BuiltInUnit.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FVal <> stbiuThousands then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_BuiltInUnit.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_BuiltInUnit.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FVal <> stbiuThousands then 
    AWriter.AddAttribute('val',StrTST_BuiltInUnit[Integer(FVal)]);
end;

procedure TCT_BuiltInUnit.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then 
    FVal := TST_BuiltInUnit(StrToEnum('stbiu' + AAttributes.Values[0]))
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_BuiltInUnit.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := stbiuThousands;
end;

destructor TCT_BuiltInUnit.Destroy;
begin
end;

procedure TCT_BuiltInUnit.Clear;
begin
  FAssigneds := [];
  FVal := stbiuThousands;
end;

procedure TCT_BuiltInUnit.Assign(AItem: TCT_BuiltInUnit);
begin
end;

procedure TCT_BuiltInUnit.CopyTo(AItem: TCT_BuiltInUnit);
begin
end;

{ TCT_DispUnitsLbl }

function  TCT_DispUnitsLbl.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FLayout <> Nil then 
    Inc(ElemsAssigned,FLayout.CheckAssigned);
  if FTx <> Nil then 
    Inc(ElemsAssigned,FTx.CheckAssigned);
  if FSpPr <> Nil then 
    Inc(ElemsAssigned,FSpPr.CheckAssigned);
  if FTxPr <> Nil then 
    Inc(ElemsAssigned,FTxPr.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_DispUnitsLbl.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $0000033B: begin
      if FLayout = Nil then 
        FLayout := TCT_Layout.Create(FOwner);
      Result := FLayout;
    end;
    $00000189: begin
      if FTx = Nil then 
        FTx := TCT_Tx.Create(FOwner);
      Result := FTx;
    end;
    $00000242: begin
      if FSpPr = Nil then 
        FSpPr := TCT_ShapeProperties.Create(FOwner);
      Result := FSpPr;
    end;
    $0000024B: begin
      if FTxPr = Nil then 
        FTxPr := TCT_TextBody.Create(FOwner);
      Result := FTxPr;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_DispUnitsLbl.Write(AWriter: TXpgWriteXML);
begin
  if (FLayout <> Nil) and FLayout.Assigned then 
    if xaElements in FLayout.Assigneds then
    begin
      AWriter.BeginTag('c:layout');
      FLayout.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:layout');
  if (FTx <> Nil) and FTx.Assigned then 
    if xaElements in FTx.Assigneds then
    begin
      AWriter.BeginTag('c:tx');
      FTx.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:tx');
  if (FSpPr <> Nil) and FSpPr.Assigned then 
  begin
    FSpPr.WriteAttributes(AWriter);
    if xaElements in FSpPr.Assigneds then
    begin
      AWriter.BeginTag('c:spPr');
      FSpPr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:spPr');
  end;
  if (FTxPr <> Nil) and FTxPr.Assigned then 
    if xaElements in FTxPr.Assigneds then
    begin
      AWriter.BeginTag('c:txPr');
      FTxPr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:txPr');
end;

constructor TCT_DispUnitsLbl.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 4;
  FAttributeCount := 0;
end;

destructor TCT_DispUnitsLbl.Destroy;
begin
  if FLayout <> Nil then 
    FLayout.Free;
  if FTx <> Nil then 
    FTx.Free;
  if FSpPr <> Nil then 
    FSpPr.Free;
  if FTxPr <> Nil then 
    FTxPr.Free;
end;

procedure TCT_DispUnitsLbl.Clear;
begin
  FAssigneds := [];
  if FLayout <> Nil then 
    FreeAndNil(FLayout);
  if FTx <> Nil then 
    FreeAndNil(FTx);
  if FSpPr <> Nil then 
    FreeAndNil(FSpPr);
  if FTxPr <> Nil then 
    FreeAndNil(FTxPr);
end;

procedure TCT_DispUnitsLbl.Assign(AItem: TCT_DispUnitsLbl);
begin
end;

procedure TCT_DispUnitsLbl.CopyTo(AItem: TCT_DispUnitsLbl);
begin
end;

function  TCT_DispUnitsLbl.Create_Layout: TCT_Layout;
begin
  Result := TCT_Layout.Create(FOwner);
  FLayout := Result;
end;

function  TCT_DispUnitsLbl.Create_Tx: TCT_Tx;
begin
  Result := TCT_Tx.Create(FOwner);
  FTx := Result;
end;

function  TCT_DispUnitsLbl.Create_SpPr: TCT_ShapeProperties;
begin
  Result := TCT_ShapeProperties.Create(FOwner);
  FSpPr := Result;
end;

function  TCT_DispUnitsLbl.Create_TxPr: TCT_TextBody;
begin
  Result := TCT_TextBody.Create(FOwner);
  FTxPr := Result;
end;

{ TEG_AreaChartShared }

function  TEG_AreaChartShared.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FGrouping <> Nil then 
    Inc(ElemsAssigned,FGrouping.CheckAssigned);
  if FVaryColors <> Nil then 
    Inc(ElemsAssigned,FVaryColors.CheckAssigned);
  if FSer <> Nil then 
    Inc(ElemsAssigned,FSer.CheckAssigned);
  if FDLbls <> Nil then 
    Inc(ElemsAssigned,FDLbls.CheckAssigned);
  if FDropLines <> Nil then 
    Inc(ElemsAssigned,FDropLines.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TEG_AreaChartShared.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $00000408: begin
      if FGrouping = Nil then 
        FGrouping := TCT_Grouping.Create(FOwner);
      Result := FGrouping;
    end;
    $000004D1: begin
      if FVaryColors = Nil then 
        FVaryColors := TCT_Boolean.Create(FOwner);
      Result := FVaryColors;
    end;
    $000001E7: begin
      if FSer = Nil then 
        FSer := TCT_AreaSer.Create(FOwner);
      Result := FSer;
    end;
    $0000028E: begin
      if FDLbls = Nil then 
        FDLbls := TCT_DLbls.Create(FOwner);
      Result := FDLbls;
    end;
    $0000044D: begin
      if FDropLines = Nil then 
        FDropLines := TCT_ChartLines.Create(FOwner);
      Result := FDropLines;
    end;
    else 
    begin
      Result := Nil;
      Exit;
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TEG_AreaChartShared.Write(AWriter: TXpgWriteXML);
begin
  if (FGrouping <> Nil) and FGrouping.Assigned then 
  begin
    FGrouping.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:grouping');
  end;
  if (FVaryColors <> Nil) and FVaryColors.Assigned then 
  begin
    FVaryColors.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:varyColors');
  end;
  if (FSer <> Nil) and FSer.Assigned then 
    if xaElements in FSer.Assigneds then
    begin
      AWriter.BeginTag('c:ser');
      FSer.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:ser');
  if (FDLbls <> Nil) and FDLbls.Assigned then 
    if xaElements in FDLbls.Assigneds then
    begin
      AWriter.BeginTag('c:dLbls');
      FDLbls.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:dLbls');
  if (FDropLines <> Nil) and FDropLines.Assigned then 
    if xaElements in FDropLines.Assigneds then
    begin
      AWriter.BeginTag('c:dropLines');
      FDropLines.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:dropLines');
end;

constructor TEG_AreaChartShared.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 5;
  FAttributeCount := 0;
end;

destructor TEG_AreaChartShared.Destroy;
begin
  if FGrouping <> Nil then 
    FGrouping.Free;
  if FVaryColors <> Nil then 
    FVaryColors.Free;
  if FSer <> Nil then 
    FSer.Free;
  if FDLbls <> Nil then 
    FDLbls.Free;
  if FDropLines <> Nil then 
    FDropLines.Free;
end;

procedure TEG_AreaChartShared.Clear;
begin
  FAssigneds := [];
  if FGrouping <> Nil then 
    FreeAndNil(FGrouping);
  if FVaryColors <> Nil then 
    FreeAndNil(FVaryColors);
  if FSer <> Nil then 
    FreeAndNil(FSer);
  if FDLbls <> Nil then 
    FreeAndNil(FDLbls);
  if FDropLines <> Nil then 
    FreeAndNil(FDropLines);
end;

procedure TEG_AreaChartShared.Assign(AItem: TEG_AreaChartShared);
begin
end;

procedure TEG_AreaChartShared.CopyTo(AItem: TEG_AreaChartShared);
begin
end;

function  TEG_AreaChartShared.Create_Grouping: TCT_Grouping;
begin
  Result := TCT_Grouping.Create(FOwner);
  FGrouping := Result;
end;

function  TEG_AreaChartShared.Create_VaryColors: TCT_Boolean;
begin
  Result := TCT_Boolean.Create(FOwner);
  FVaryColors := Result;
end;

function  TEG_AreaChartShared.Create_Ser: TCT_AreaSer;
begin
  Result := TCT_AreaSer.Create(FOwner);
  FSer := Result;
end;

function  TEG_AreaChartShared.Create_DLbls: TCT_DLbls;
begin
  Result := TCT_DLbls.Create(FOwner);
  FDLbls := Result;
end;

function  TEG_AreaChartShared.Create_DropLines: TCT_ChartLines;
begin
  Result := TCT_ChartLines.Create(FOwner);
  FDropLines := Result;
end;

{ TEG_LineChartShared }

function  TEG_LineChartShared.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FGrouping <> Nil then 
    Inc(ElemsAssigned,FGrouping.CheckAssigned);
  if FVaryColors <> Nil then 
    Inc(ElemsAssigned,FVaryColors.CheckAssigned);
  if FSer <> Nil then
    Inc(ElemsAssigned,FSer.CheckAssigned);
  if FDLbls <> Nil then 
    Inc(ElemsAssigned,FDLbls.CheckAssigned);
  if FDropLines <> Nil then 
    Inc(ElemsAssigned,FDropLines.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TEG_LineChartShared.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $00000408: begin
      if FGrouping = Nil then 
        FGrouping := TCT_Grouping.Create(FOwner);
      Result := FGrouping;
    end;
    $000004D1: begin
      if FVaryColors = Nil then 
        FVaryColors := TCT_Boolean.Create(FOwner);
      Result := FVaryColors;
    end;
    $000001E7: begin
      if FSer = Nil then 
        FSer := TCT_LineSerXpgList.Create(FOwner);
      Result := FSer.Add;
    end;
    $0000028E: begin
      if FDLbls = Nil then 
        FDLbls := TCT_DLbls.Create(FOwner);
      Result := FDLbls;
    end;
    $0000044D: begin
      if FDropLines = Nil then 
        FDropLines := TCT_ChartLines.Create(FOwner);
      Result := FDropLines;
    end;
    else 
    begin
      Result := Nil;
      Exit;
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TEG_LineChartShared.Write(AWriter: TXpgWriteXML);
begin
  if FGrouping <> Nil then begin
    FGrouping.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:grouping');
  end;
  if (FVaryColors <> Nil) and FVaryColors.Assigned then 
  begin
    FVaryColors.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:varyColors');
  end;

  if (FSer <> Nil) and FSer.FAssigned then
    FSer.Write(AWriter,'c:ser');

  if (FDLbls <> Nil) and FDLbls.Assigned then
    if xaElements in FDLbls.Assigneds then
    begin
      AWriter.BeginTag('c:dLbls');
      FDLbls.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:dLbls');
  if (FDropLines <> Nil) and FDropLines.Assigned then 
    if xaElements in FDropLines.Assigneds then
    begin
      AWriter.BeginTag('c:dropLines');
      FDropLines.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:dropLines');
end;

constructor TEG_LineChartShared.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 5;
  FAttributeCount := 0;
end;

destructor TEG_LineChartShared.Destroy;
begin
  if FGrouping <> Nil then 
    FGrouping.Free;
  if FVaryColors <> Nil then 
    FVaryColors.Free;
  if FSer <> Nil then 
    FSer.Free;
  if FDLbls <> Nil then 
    FDLbls.Free;
  if FDropLines <> Nil then 
    FDropLines.Free;
end;

procedure TEG_LineChartShared.Clear;
begin
  FAssigneds := [];
  if FGrouping <> Nil then 
    FreeAndNil(FGrouping);
  if FVaryColors <> Nil then 
    FreeAndNil(FVaryColors);
  if FSer <> Nil then 
    FreeAndNil(FSer);
  if FDLbls <> Nil then 
    FreeAndNil(FDLbls);
  if FDropLines <> Nil then 
    FreeAndNil(FDropLines);
end;

procedure TEG_LineChartShared.Assign(AItem: TEG_LineChartShared);
begin
end;

procedure TEG_LineChartShared.CopyTo(AItem: TEG_LineChartShared);
begin
end;

function  TEG_LineChartShared.Create_Grouping: TCT_Grouping;
begin
  Result := TCT_Grouping.Create(FOwner);
  FGrouping := Result;
end;

function  TEG_LineChartShared.Create_VaryColors: TCT_Boolean;
begin
  Result := TCT_Boolean.Create(FOwner);
  FVaryColors := Result;
end;

function  TEG_LineChartShared.Create_Ser: TCT_LineSerXpgList;
begin
  Result := TCT_LineSerXpgList.Create(FOwner);
  FSer := Result;
end;

function  TEG_LineChartShared.Create_DLbls: TCT_DLbls;
begin
  Result := TCT_DLbls.Create(FOwner);
  FDLbls := Result;
end;

function  TEG_LineChartShared.Create_DropLines: TCT_ChartLines;
begin
  Result := TCT_ChartLines.Create(FOwner);
  FDropLines := Result;
end;

{ TCT_UpDownBars }

function  TCT_UpDownBars.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FGapWidth <> Nil then 
    Inc(ElemsAssigned,FGapWidth.CheckAssigned);
  if FUpBars <> Nil then 
    Inc(ElemsAssigned,FUpBars.CheckAssigned);
  if FDownBars <> Nil then 
    Inc(ElemsAssigned,FDownBars.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_UpDownBars.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $000003D5: begin
      if FGapWidth = Nil then 
        FGapWidth := TCT_GapAmount.Create(FOwner);
      Result := FGapWidth;
    end;
    $0000030A: begin
      if FUpBars = Nil then 
        FUpBars := TCT_UpDownBar.Create(FOwner);
      Result := FUpBars;
    end;
    $000003DD: begin
      if FDownBars = Nil then 
        FDownBars := TCT_UpDownBar.Create(FOwner);
      Result := FDownBars;
    end;
    $00000321: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_UpDownBars.Write(AWriter: TXpgWriteXML);
begin
  if (FGapWidth <> Nil) and FGapWidth.Assigned then 
  begin
    FGapWidth.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:gapWidth');
  end;
  if (FUpBars <> Nil) and FUpBars.Assigned then 
    if xaElements in FUpBars.Assigneds then
    begin
      AWriter.BeginTag('c:upBars');
      FUpBars.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:upBars');
  if (FDownBars <> Nil) and FDownBars.Assigned then 
    if xaElements in FDownBars.Assigneds then
    begin
      AWriter.BeginTag('c:downBars');
      FDownBars.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:downBars');
{$ifdef XLS_WRITE_EXTLST}
  if (FExtLst <> Nil) and FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then
    begin
      AWriter.BeginTag('c:extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:extLst');
{$endif}
end;

constructor TCT_UpDownBars.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 4;
  FAttributeCount := 0;
end;

destructor TCT_UpDownBars.Destroy;
begin
  if FGapWidth <> Nil then 
    FGapWidth.Free;
  if FUpBars <> Nil then 
    FUpBars.Free;
  if FDownBars <> Nil then 
    FDownBars.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_UpDownBars.Clear;
begin
  FAssigneds := [];
  if FGapWidth <> Nil then 
    FreeAndNil(FGapWidth);
  if FUpBars <> Nil then 
    FreeAndNil(FUpBars);
  if FDownBars <> Nil then 
    FreeAndNil(FDownBars);
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
end;

procedure TCT_UpDownBars.Assign(AItem: TCT_UpDownBars);
begin
end;

procedure TCT_UpDownBars.CopyTo(AItem: TCT_UpDownBars);
begin
end;

function  TCT_UpDownBars.Create_GapWidth: TCT_GapAmount;
begin
  Result := TCT_GapAmount.Create(FOwner);
  FGapWidth := Result;
end;

function  TCT_UpDownBars.Create_UpBars: TCT_UpDownBar;
begin
  Result := TCT_UpDownBar.Create(FOwner);
  FUpBars := Result;
end;

function  TCT_UpDownBars.Create_DownBars: TCT_UpDownBar;
begin
  Result := TCT_UpDownBar.Create(FOwner);
  FDownBars := Result;
end;

function  TCT_UpDownBars.Create_ExtLst: TCT_ExtensionList;
begin
  Result := TCT_ExtensionList.Create(FOwner);
  FExtLst := Result;
end;

{ TCT_RadarStyle }

function  TCT_RadarStyle.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FVal <> strsStandard then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_RadarStyle.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_RadarStyle.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FVal <> strsStandard then 
    AWriter.AddAttribute('val',StrTST_RadarStyle[Integer(FVal)]);
end;

procedure TCT_RadarStyle.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then 
    FVal := TST_RadarStyle(StrToEnum('strs' + AAttributes.Values[0]))
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_RadarStyle.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := strsStandard;
end;

destructor TCT_RadarStyle.Destroy;
begin
end;

procedure TCT_RadarStyle.Clear;
begin
  FAssigneds := [];
  FVal := strsStandard;
end;

procedure TCT_RadarStyle.Assign(AItem: TCT_RadarStyle);
begin
end;

procedure TCT_RadarStyle.CopyTo(AItem: TCT_RadarStyle);
begin
end;

{ TCT_RadarSer }

function  TCT_RadarSer.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FEG_SerShared.CheckAssigned);
  if FMarker <> Nil then 
    Inc(ElemsAssigned,FMarker.CheckAssigned);
  if FDPtXpgList <> Nil then 
    Inc(ElemsAssigned,FDPtXpgList.CheckAssigned);
  if FDLbls <> Nil then 
    Inc(ElemsAssigned,FDLbls.CheckAssigned);
  if FCat <> Nil then 
    Inc(ElemsAssigned,FCat.CheckAssigned);
  if FVal <> Nil then 
    Inc(ElemsAssigned,FVal.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_RadarSer.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $0000031F: begin
      if FMarker = Nil then 
        FMarker := TCT_Marker.Create(FOwner);
      Result := FMarker;
    end;
    $000001C5: begin
      if FDPtXpgList = Nil then 
        FDPtXpgList := TCT_DPtXpgList.Create(FOwner);
      Result := FDPtXpgList.Add;
    end;
    $0000028E: begin
      if FDLbls = Nil then 
        FDLbls := TCT_DLbls.Create(FOwner);
      Result := FDLbls;
    end;
    $000001D5: begin
      if FCat = Nil then 
        FCat := TCT_AxDataSource.Create(FOwner);
      Result := FCat;
    end;
    $000001E0: begin
      if FVal = Nil then 
        FVal := TCT_NumDataSource.Create(FOwner);
      Result := FVal;
    end;
    $00000321: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
    begin
      Result := FEG_SerShared.HandleElement(AReader);
      if Result = Nil then 
        FOwner.Errors.Error(xemUnknownElement,AReader.QName);
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_RadarSer.Write(AWriter: TXpgWriteXML);
begin
  FEG_SerShared.Write(AWriter);
  if (FMarker <> Nil) and FMarker.Assigned then 
    if xaElements in FMarker.Assigneds then
    begin
      AWriter.BeginTag('c:marker');
      FMarker.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:marker');
  if FDPtXpgList <> Nil then 
    FDPtXpgList.Write(AWriter,'c:dPt');
  if (FDLbls <> Nil) and FDLbls.Assigned then 
    if xaElements in FDLbls.Assigneds then
    begin
      AWriter.BeginTag('c:dLbls');
      FDLbls.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:dLbls');
  if (FCat <> Nil) and FCat.Assigned then 
    if xaElements in FCat.Assigneds then
    begin
      AWriter.BeginTag('c:cat');
      FCat.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:cat');
  if (FVal <> Nil) and FVal.Assigned then 
    if xaElements in FVal.Assigneds then
    begin
      AWriter.BeginTag('c:val');
      FVal.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:val');
{$ifdef XLS_WRITE_EXTLST}
  if (FExtLst <> Nil) and FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then
    begin
      AWriter.BeginTag('c:extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:extLst');
{$endif}
end;

constructor TCT_RadarSer.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 7;
  FAttributeCount := 0;
  FEG_SerShared := TEG_SerShared.Create(FOwner);
end;

destructor TCT_RadarSer.Destroy;
begin
  FEG_SerShared.Free;
  if FMarker <> Nil then 
    FMarker.Free;
  if FDPtXpgList <> Nil then 
    FDPtXpgList.Free;
  if FDLbls <> Nil then 
    FDLbls.Free;
  if FCat <> Nil then 
    FCat.Free;
  if FVal <> Nil then 
    FVal.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_RadarSer.Clear;
begin
  FAssigneds := [];
  FEG_SerShared.Clear;
  if FMarker <> Nil then 
    FreeAndNil(FMarker);
  if FDPtXpgList <> Nil then 
    FreeAndNil(FDPtXpgList);
  if FDLbls <> Nil then 
    FreeAndNil(FDLbls);
  if FCat <> Nil then 
    FreeAndNil(FCat);
  if FVal <> Nil then 
    FreeAndNil(FVal);
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
end;

procedure TCT_RadarSer.Assign(AItem: TCT_RadarSer);
begin
end;

procedure TCT_RadarSer.CopyTo(AItem: TCT_RadarSer);
begin
end;

function  TCT_RadarSer.Create_Marker: TCT_Marker;
begin
  Result := TCT_Marker.Create(FOwner);
  FMarker := Result;
end;

function  TCT_RadarSer.Create_DPtXpgList: TCT_DPtXpgList;
begin
  Result := TCT_DPtXpgList.Create(FOwner);
  FDPtXpgList := Result;
end;

function  TCT_RadarSer.Create_DLbls: TCT_DLbls;
begin
  Result := TCT_DLbls.Create(FOwner);
  FDLbls := Result;
end;

function  TCT_RadarSer.Create_Cat: TCT_AxDataSource;
begin
  Result := TCT_AxDataSource.Create(FOwner);
  FCat := Result;
end;

function  TCT_RadarSer.Create_Val: TCT_NumDataSource;
begin
  Result := TCT_NumDataSource.Create(FOwner);
  FVal := Result;
end;

function  TCT_RadarSer.Create_ExtLst: TCT_ExtensionList;
begin
  Result := TCT_ExtensionList.Create(FOwner);
  FExtLst := Result;
end;

{ TCT_RadarSerXpgList }

function  TCT_RadarSerXpgList.GetItems(Index: integer): TCT_RadarSer;
begin
  Result := TCT_RadarSer(inherited Items[Index]);
end;

function  TCT_RadarSerXpgList.Add: TCT_RadarSer;
begin
  Result := TCT_RadarSer.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_RadarSerXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_RadarSerXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
    if xaElements in Items[i].Assigneds then
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
end;

procedure TCT_RadarSerXpgList.Assign(AItem: TCT_RadarSerXpgList);
begin
end;

procedure TCT_RadarSerXpgList.CopyTo(AItem: TCT_RadarSerXpgList);
begin
end;

{ TCT_ScatterStyle }

function  TCT_ScatterStyle.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FVal <> stssMarker then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_ScatterStyle.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_ScatterStyle.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FVal <> stssMarker then 
    AWriter.AddAttribute('val',StrTST_ScatterStyle[Integer(FVal)]);
end;

procedure TCT_ScatterStyle.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then 
    FVal := TST_ScatterStyle(StrToEnum('stss' + AAttributes.Values[0]))
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_ScatterStyle.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := stssMarker;
end;

destructor TCT_ScatterStyle.Destroy;
begin
end;

procedure TCT_ScatterStyle.Clear;
begin
  FAssigneds := [];
  FVal := stssMarker;
end;

procedure TCT_ScatterStyle.Assign(AItem: TCT_ScatterStyle);
begin
end;

procedure TCT_ScatterStyle.CopyTo(AItem: TCT_ScatterStyle);
begin
end;

{ TCT_ScatterSer }

function  TCT_ScatterSer.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FEG_SerShared.CheckAssigned);
  if FMarker <> Nil then 
    Inc(ElemsAssigned,FMarker.CheckAssigned);
  if FDPtXpgList <> Nil then 
    Inc(ElemsAssigned,FDPtXpgList.CheckAssigned);
  if FDLbls <> Nil then 
    Inc(ElemsAssigned,FDLbls.CheckAssigned);
  if FTrendlineXpgList <> Nil then 
    Inc(ElemsAssigned,FTrendlineXpgList.CheckAssigned);
  if FErrBarsXpgList <> Nil then 
    Inc(ElemsAssigned,FErrBarsXpgList.CheckAssigned);
  if FXVal <> Nil then 
    Inc(ElemsAssigned,FXVal.CheckAssigned);
  if FYVal <> Nil then 
    Inc(ElemsAssigned,FYVal.CheckAssigned);
  if FSmooth <> Nil then 
    Inc(ElemsAssigned,FSmooth.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_ScatterSer.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $0000031F: begin
      if FMarker = Nil then 
        FMarker := TCT_Marker.Create(FOwner);
      Result := FMarker;
    end;
    $000001C5: begin
      if FDPtXpgList = Nil then 
        FDPtXpgList := TCT_DPtXpgList.Create(FOwner);
      Result := FDPtXpgList.Add;
    end;
    $0000028E: begin
      if FDLbls = Nil then 
        FDLbls := TCT_DLbls.Create(FOwner);
      Result := FDLbls;
    end;
    $00000462: begin
      if FTrendlineXpgList = Nil then 
        FTrendlineXpgList := TCT_TrendlineXpgList.Create(FOwner);
      Result := FTrendlineXpgList.Add;
    end;
    $0000036E: begin
      if FErrBarsXpgList = Nil then 
        FErrBarsXpgList := TCT_ErrBarsXpgList.Create(FOwner);
      Result := FErrBarsXpgList.Add;
    end;
    $00000238: begin
      if FXVal = Nil then 
        FXVal := TCT_AxDataSource.Create(FOwner);
      Result := FXVal;
    end;
    $00000239: begin
      if FYVal = Nil then 
        FYVal := TCT_NumDataSource.Create(FOwner);
      Result := FYVal;
    end;
    $00000337: begin
      if FSmooth = Nil then 
        FSmooth := TCT_Boolean.Create(FOwner);
      Result := FSmooth;
    end;
    $00000321: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
    begin
      Result := FEG_SerShared.HandleElement(AReader);
      if Result = Nil then 
        FOwner.Errors.Error(xemUnknownElement,AReader.QName);
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_ScatterSer.Write(AWriter: TXpgWriteXML);
begin
  FEG_SerShared.Write(AWriter);
  if (FMarker <> Nil) and FMarker.Assigned then 
    if xaElements in FMarker.Assigneds then
    begin
      AWriter.BeginTag('c:marker');
      FMarker.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:marker');
  if FDPtXpgList <> Nil then 
    FDPtXpgList.Write(AWriter,'c:dPt');
  if (FDLbls <> Nil) and FDLbls.Assigned then 
    if xaElements in FDLbls.Assigneds then
    begin
      AWriter.BeginTag('c:dLbls');
      FDLbls.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:dLbls');
  if FTrendlineXpgList <> Nil then 
    FTrendlineXpgList.Write(AWriter,'c:trendline');
  if FErrBarsXpgList <> Nil then 
    FErrBarsXpgList.Write(AWriter,'c:errBars');
  if (FXVal <> Nil) and FXVal.Assigned then 
    if xaElements in FXVal.Assigneds then
    begin
      AWriter.BeginTag('c:xVal');
      FXVal.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:xVal');
  if (FYVal <> Nil) and FYVal.Assigned then 
    if xaElements in FYVal.Assigneds then
    begin
      AWriter.BeginTag('c:yVal');
      FYVal.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:yVal');
  if (FSmooth <> Nil) and FSmooth.Assigned then 
  begin
    FSmooth.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:smooth');
  end;
{$ifdef XLS_WRITE_EXTLST}
  if (FExtLst <> Nil) and FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then
    begin
      AWriter.BeginTag('c:extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:extLst');
{$endif}
end;

constructor TCT_ScatterSer.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 10;
  FAttributeCount := 0;
  FEG_SerShared := TEG_SerShared.Create(FOwner);
end;

destructor TCT_ScatterSer.Destroy;
begin
  FEG_SerShared.Free;
  if FMarker <> Nil then 
    FMarker.Free;
  if FDPtXpgList <> Nil then 
    FDPtXpgList.Free;
  if FDLbls <> Nil then 
    FDLbls.Free;
  if FTrendlineXpgList <> Nil then 
    FTrendlineXpgList.Free;
  if FErrBarsXpgList <> Nil then 
    FErrBarsXpgList.Free;
  if FXVal <> Nil then 
    FXVal.Free;
  if FYVal <> Nil then 
    FYVal.Free;
  if FSmooth <> Nil then 
    FSmooth.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_ScatterSer.Clear;
begin
  FAssigneds := [];
  FEG_SerShared.Clear;
  if FMarker <> Nil then 
    FreeAndNil(FMarker);
  if FDPtXpgList <> Nil then 
    FreeAndNil(FDPtXpgList);
  if FDLbls <> Nil then 
    FreeAndNil(FDLbls);
  if FTrendlineXpgList <> Nil then 
    FreeAndNil(FTrendlineXpgList);
  if FErrBarsXpgList <> Nil then 
    FreeAndNil(FErrBarsXpgList);
  if FXVal <> Nil then 
    FreeAndNil(FXVal);
  if FYVal <> Nil then 
    FreeAndNil(FYVal);
  if FSmooth <> Nil then 
    FreeAndNil(FSmooth);
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
end;

procedure TCT_ScatterSer.Assign(AItem: TCT_ScatterSer);
begin
end;

procedure TCT_ScatterSer.CopyTo(AItem: TCT_ScatterSer);
begin
end;

function  TCT_ScatterSer.Create_Marker: TCT_Marker;
begin
  Result := TCT_Marker.Create(FOwner);
  FMarker := Result;
end;

function  TCT_ScatterSer.Create_DPtXpgList: TCT_DPtXpgList;
begin
  Result := TCT_DPtXpgList.Create(FOwner);
  FDPtXpgList := Result;
end;

function  TCT_ScatterSer.Create_DLbls: TCT_DLbls;
begin
  Result := TCT_DLbls.Create(FOwner);
  FDLbls := Result;
end;

function  TCT_ScatterSer.Create_TrendlineXpgList: TCT_TrendlineXpgList;
begin
  Result := TCT_TrendlineXpgList.Create(FOwner);
  FTrendlineXpgList := Result;
end;

function  TCT_ScatterSer.Create_ErrBarsXpgList: TCT_ErrBarsXpgList;
begin
  Result := TCT_ErrBarsXpgList.Create(FOwner);
  FErrBarsXpgList := Result;
end;

function  TCT_ScatterSer.Create_XVal: TCT_AxDataSource;
begin
  Result := TCT_AxDataSource.Create(FOwner);
  FXVal := Result;
end;

function  TCT_ScatterSer.Create_YVal: TCT_NumDataSource;
begin
  Result := TCT_NumDataSource.Create(FOwner);
  FYVal := Result;
end;

function  TCT_ScatterSer.Create_Smooth: TCT_Boolean;
begin
  Result := TCT_Boolean.Create(FOwner);
  FSmooth := Result;
end;

function  TCT_ScatterSer.Create_ExtLst: TCT_ExtensionList;
begin
  Result := TCT_ExtensionList.Create(FOwner);
  FExtLst := Result;
end;

{ TCT_ScatterSerXpgList }

function  TCT_ScatterSerXpgList.GetItems(Index: integer): TCT_ScatterSer;
begin
  Result := TCT_ScatterSer(inherited Items[Index]);
end;

function  TCT_ScatterSerXpgList.Add: TCT_ScatterSer;
begin
  Result := TCT_ScatterSer.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_ScatterSerXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_ScatterSerXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
    if xaElements in Items[i].Assigneds then
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
end;

procedure TCT_ScatterSerXpgList.Assign(AItem: TCT_ScatterSerXpgList);
begin
end;

procedure TCT_ScatterSerXpgList.CopyTo(AItem: TCT_ScatterSerXpgList);
begin
end;

{ TEG_PieChartShared }

function  TEG_PieChartShared.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FVaryColors <> Nil then 
    Inc(ElemsAssigned,FVaryColors.CheckAssigned);
  if FSer <> Nil then 
    Inc(ElemsAssigned,FSer.CheckAssigned);
  if FDLbls <> Nil then 
    Inc(ElemsAssigned,FDLbls.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TEG_PieChartShared.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $000004D1: begin
      if FVaryColors = Nil then 
        FVaryColors := TCT_Boolean.Create(FOwner);
      Result := FVaryColors;
    end;
    $000001E7: begin
      if FSer = Nil then 
        FSer := TCT_PieSer.Create(FOwner);
      Result := FSer;
    end;
    $0000028E: begin
      if FDLbls = Nil then 
        FDLbls := TCT_DLbls.Create(FOwner);
      Result := FDLbls;
    end;
    else 
    begin
      Result := Nil;
      Exit;
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TEG_PieChartShared.Write(AWriter: TXpgWriteXML);
begin
  if (FVaryColors <> Nil) and FVaryColors.Assigned then 
  begin
    FVaryColors.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:varyColors');
  end;
  if (FSer <> Nil) and FSer.Assigned then 
    if xaElements in FSer.Assigneds then
    begin
      AWriter.BeginTag('c:ser');
      FSer.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:ser');
  if (FDLbls <> Nil) and FDLbls.Assigned then 
    if xaElements in FDLbls.Assigneds then
    begin
      AWriter.BeginTag('c:dLbls');
      FDLbls.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:dLbls');
end;

constructor TEG_PieChartShared.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 3;
  FAttributeCount := 0;
end;

destructor TEG_PieChartShared.Destroy;
begin
  if FVaryColors <> Nil then 
    FVaryColors.Free;
  if FSer <> Nil then 
    FSer.Free;
  if FDLbls <> Nil then 
    FDLbls.Free;
end;

procedure TEG_PieChartShared.Clear;
begin
  FAssigneds := [];
  if FVaryColors <> Nil then 
    FreeAndNil(FVaryColors);
  if FSer <> Nil then 
    FreeAndNil(FSer);
  if FDLbls <> Nil then 
    FreeAndNil(FDLbls);
end;

procedure TEG_PieChartShared.Assign(AItem: TEG_PieChartShared);
begin
end;

procedure TEG_PieChartShared.CopyTo(AItem: TEG_PieChartShared);
begin
end;

function  TEG_PieChartShared.Create_VaryColors: TCT_Boolean;
begin
  Result := TCT_Boolean.Create(FOwner);
  FVaryColors := Result;
end;

function  TEG_PieChartShared.Create_Ser: TCT_PieSer;
begin
  Result := TCT_PieSer.Create(FOwner);
  FSer := Result;
end;

function  TEG_PieChartShared.Create_DLbls: TCT_DLbls;
begin
  Result := TCT_DLbls.Create(FOwner);
  FDLbls := Result;
end;

{ TCT_FirstSliceAng }

function  TCT_FirstSliceAng.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FVal <> 0 then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_FirstSliceAng.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_FirstSliceAng.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FVal <> 0 then 
    AWriter.AddAttribute('val',XmlIntToStr(FVal));
end;

procedure TCT_FirstSliceAng.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then 
    FVal := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_FirstSliceAng.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := 0;
end;

destructor TCT_FirstSliceAng.Destroy;
begin
end;

procedure TCT_FirstSliceAng.Clear;
begin
  FAssigneds := [];
  FVal := 0;
end;

procedure TCT_FirstSliceAng.Assign(AItem: TCT_FirstSliceAng);
begin
end;

procedure TCT_FirstSliceAng.CopyTo(AItem: TCT_FirstSliceAng);
begin
end;

{ TCT_HoleSize }

function  TCT_HoleSize.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FVal <> 10 then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_HoleSize.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_HoleSize.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FVal <> 10 then 
    AWriter.AddAttribute('val',XmlIntToStr(FVal));
end;

procedure TCT_HoleSize.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then 
    FVal := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_HoleSize.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := 10;
end;

destructor TCT_HoleSize.Destroy;
begin
end;

procedure TCT_HoleSize.Clear;
begin
  FAssigneds := [];
  FVal := 10;
end;

procedure TCT_HoleSize.Assign(AItem: TCT_HoleSize);
begin
end;

procedure TCT_HoleSize.CopyTo(AItem: TCT_HoleSize);
begin
end;

{ TEG_BarChartShared }

function  TEG_BarChartShared.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FBarDir <> Nil then 
    Inc(ElemsAssigned,FBarDir.CheckAssigned);
  if FGrouping <> Nil then 
    Inc(ElemsAssigned,FGrouping.CheckAssigned);
  if FVaryColors <> Nil then 
    Inc(ElemsAssigned,FVaryColors.CheckAssigned);
  if FSeries.Count > 0 then begin
    FSeries.CheckAssigned;
    Inc(ElemsAssigned);
  end;
  if FDLbls <> Nil then
    Inc(ElemsAssigned,FDLbls.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TEG_BarChartShared.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $000002F1: begin
      if FBarDir = Nil then 
        FBarDir := TCT_BarDir.Create(FOwner);
      Result := FBarDir;
    end;
    $00000408: begin
      if FGrouping = Nil then 
        FGrouping := TCT_BarGrouping.Create(FOwner);
      Result := FGrouping;
    end;
    $000004D1: begin
      if FVaryColors = Nil then 
        FVaryColors := TCT_Boolean.Create(FOwner);
      Result := FVaryColors;
    end;
    $000001E7: begin
      Result := FSeries.Add(TCT_BarSer.Create(FOwner));
    end;
    $0000028E: begin
      if FDLbls = Nil then 
        FDLbls := TCT_DLbls.Create(FOwner);
      Result := FDLbls;
    end;
    else 
    begin
      Result := Nil;
      Exit;
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TEG_BarChartShared.Write(AWriter: TXpgWriteXML);
var
  i: integer;
begin
  if (FBarDir <> Nil) and FBarDir.Assigned then
    FBarDir.WriteAttributes(AWriter);
  AWriter.SimpleTag('c:barDir');

  if (FGrouping <> Nil) and FGrouping.Assigned then
    FGrouping.WriteAttributes(AWriter);
  AWriter.SimpleTag('c:grouping');

  if (FVaryColors <> Nil) and FVaryColors.Assigned then
  begin
    FVaryColors.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:varyColors');
  end;
  for i := 0 to FSeries.Count - 1 do begin
     AWriter.BeginTag('c:ser');
     TCT_BarSer(FSeries[i]).Write(AWriter);
     AWriter.EndTag;
  end;
  if (FDLbls <> Nil) and FDLbls.Assigned then
    if xaElements in FDLbls.Assigneds then
    begin
      AWriter.BeginTag('c:dLbls');
      FDLbls.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:dLbls');
end;

constructor TEG_BarChartShared.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 5;
  FAttributeCount := 0;
  FSeries := TXPGBaseList.Create
end;

destructor TEG_BarChartShared.Destroy;
begin
  if FBarDir <> Nil then
    FBarDir.Free;
  if FGrouping <> Nil then
    FGrouping.Free;
  if FVaryColors <> Nil then
    FVaryColors.Free;
  FSeries.Free;
  if FDLbls <> Nil then
    FDLbls.Free;
end;

procedure TEG_BarChartShared.Clear;
begin
  FAssigneds := [];
  if FBarDir <> Nil then
    FreeAndNil(FBarDir);
  if FGrouping <> Nil then
    FreeAndNil(FGrouping);
  if FVaryColors <> Nil then
    FreeAndNil(FVaryColors);
  FSeries.Clear;
  if FDLbls <> Nil then 
    FreeAndNil(FDLbls);
end;

procedure TEG_BarChartShared.Assign(AItem: TEG_BarChartShared);
begin
end;

procedure TEG_BarChartShared.CopyTo(AItem: TEG_BarChartShared);
begin
end;

function  TEG_BarChartShared.Create_BarDir: TCT_BarDir;
begin
  Result := TCT_BarDir.Create(FOwner);
  FBarDir := Result;
end;

function  TEG_BarChartShared.Create_Grouping: TCT_BarGrouping;
begin
  Result := TCT_BarGrouping.Create(FOwner);
  FGrouping := Result;
end;

function  TEG_BarChartShared.Create_VaryColors: TCT_Boolean;
begin
  Result := TCT_Boolean.Create(FOwner);
  FVaryColors := Result;
end;

function  TEG_BarChartShared.Create_DLbls: TCT_DLbls;
begin
  Result := TCT_DLbls.Create(FOwner);
  FDLbls := Result;
end;

{ TCT_Overlap }

function  TCT_Overlap.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FVal <> 0 then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_Overlap.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_Overlap.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FVal <> 0 then 
    AWriter.AddAttribute('val',XmlIntToStr(FVal));
end;

procedure TCT_Overlap.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then 
    FVal := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_Overlap.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := 0;
end;

destructor TCT_Overlap.Destroy;
begin
end;

procedure TCT_Overlap.Clear;
begin
  FAssigneds := [];
  FVal := 0;
end;

procedure TCT_Overlap.Assign(AItem: TCT_Overlap);
begin
end;

procedure TCT_Overlap.CopyTo(AItem: TCT_Overlap);
begin
end;

{ TCT_OfPieType }

function  TCT_OfPieType.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FVal <> stoptPie then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_OfPieType.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_OfPieType.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FVal <> stoptPie then 
    AWriter.AddAttribute('val',StrTST_OfPieType[Integer(FVal)]);
end;

procedure TCT_OfPieType.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then 
    FVal := TST_OfPieType(StrToEnum('stopt' + AAttributes.Values[0]))
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_OfPieType.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := stoptPie;
end;

destructor TCT_OfPieType.Destroy;
begin
end;

procedure TCT_OfPieType.Clear;
begin
  FAssigneds := [];
  FVal := stoptPie;
end;

procedure TCT_OfPieType.Assign(AItem: TCT_OfPieType);
begin
end;

procedure TCT_OfPieType.CopyTo(AItem: TCT_OfPieType);
begin
end;

{ TCT_SplitType }

function  TCT_SplitType.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FVal <> ststAuto then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_SplitType.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_SplitType.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FVal <> ststAuto then 
    AWriter.AddAttribute('val',StrTST_SplitType[Integer(FVal)]);
end;

procedure TCT_SplitType.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then 
    FVal := TST_SplitType(StrToEnum('stst' + AAttributes.Values[0]))
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_SplitType.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := ststAuto;
end;

destructor TCT_SplitType.Destroy;
begin
end;

procedure TCT_SplitType.Clear;
begin
  FAssigneds := [];
  FVal := ststAuto;
end;

procedure TCT_SplitType.Assign(AItem: TCT_SplitType);
begin
end;

procedure TCT_SplitType.CopyTo(AItem: TCT_SplitType);
begin
end;

{ TCT_CustSplit }

function  TCT_CustSplit.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FSecondPiePtXpgList <> Nil then 
    Inc(ElemsAssigned,FSecondPiePtXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_CustSplit.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'c:secondPiePt' then 
  begin
    if FSecondPiePtXpgList = Nil then 
      FSecondPiePtXpgList := TCT_UnsignedIntXpgList.Create(FOwner);
    Result := FSecondPiePtXpgList.Add;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_CustSplit.Write(AWriter: TXpgWriteXML);
begin
  if FSecondPiePtXpgList <> Nil then 
    FSecondPiePtXpgList.Write(AWriter,'c:secondPiePt');
end;

constructor TCT_CustSplit.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
end;

destructor TCT_CustSplit.Destroy;
begin
  if FSecondPiePtXpgList <> Nil then 
    FSecondPiePtXpgList.Free;
end;

procedure TCT_CustSplit.Clear;
begin
  FAssigneds := [];
  if FSecondPiePtXpgList <> Nil then 
    FreeAndNil(FSecondPiePtXpgList);
end;

procedure TCT_CustSplit.Assign(AItem: TCT_CustSplit);
begin
end;

procedure TCT_CustSplit.CopyTo(AItem: TCT_CustSplit);
begin
end;

function  TCT_CustSplit.Create_SecondPiePtXpgList: TCT_UnsignedIntXpgList;
begin
  Result := TCT_UnsignedIntXpgList.Create(FOwner);
  FSecondPiePtXpgList := Result;
end;

{ TCT_SecondPieSize }

function  TCT_SecondPieSize.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FVal <> 75 then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_SecondPieSize.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_SecondPieSize.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FVal <> 75 then 
    AWriter.AddAttribute('val',XmlIntToStr(FVal));
end;

procedure TCT_SecondPieSize.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then 
    FVal := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_SecondPieSize.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := 75;
end;

destructor TCT_SecondPieSize.Destroy;
begin
end;

procedure TCT_SecondPieSize.Clear;
begin
  FAssigneds := [];
  FVal := 75;
end;

procedure TCT_SecondPieSize.Assign(AItem: TCT_SecondPieSize);
begin
end;

procedure TCT_SecondPieSize.CopyTo(AItem: TCT_SecondPieSize);
begin
end;

{ TEG_SurfaceChartShared }

function  TEG_SurfaceChartShared.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FWireframe <> Nil then 
    Inc(ElemsAssigned,FWireframe.CheckAssigned);
  if FSer <> Nil then 
    Inc(ElemsAssigned,FSer.CheckAssigned);
  if FBandFmts <> Nil then 
    Inc(ElemsAssigned,FBandFmts.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TEG_SurfaceChartShared.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $0000045F: begin
      if FWireframe = Nil then 
        FWireframe := TCT_Boolean.Create(FOwner);
      Result := FWireframe;
    end;
    $000001E7: begin
      if FSer = Nil then 
        FSer := TCT_SurfaceSer.Create(FOwner);
      Result := FSer;
    end;
    $000003CC: begin
      if FBandFmts = Nil then 
        FBandFmts := TCT_BandFmts.Create(FOwner);
      Result := FBandFmts;
    end;
    else 
    begin
      Result := Nil;
      Exit;
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TEG_SurfaceChartShared.Write(AWriter: TXpgWriteXML);
begin
  if (FWireframe <> Nil) and FWireframe.Assigned then 
  begin
    FWireframe.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:wireframe');
  end;
  if (FSer <> Nil) and FSer.Assigned then 
    if xaElements in FSer.Assigneds then
    begin
      AWriter.BeginTag('c:ser');
      FSer.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:ser');
  if (FBandFmts <> Nil) and FBandFmts.Assigned then 
    if xaElements in FBandFmts.Assigneds then
    begin
      AWriter.BeginTag('c:bandFmts');
      FBandFmts.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:bandFmts');
end;

constructor TEG_SurfaceChartShared.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 3;
  FAttributeCount := 0;
end;

destructor TEG_SurfaceChartShared.Destroy;
begin
  if FWireframe <> Nil then 
    FWireframe.Free;
  if FSer <> Nil then 
    FSer.Free;
  if FBandFmts <> Nil then 
    FBandFmts.Free;
end;

procedure TEG_SurfaceChartShared.Clear;
begin
  FAssigneds := [];
  if FWireframe <> Nil then 
    FreeAndNil(FWireframe);
  if FSer <> Nil then 
    FreeAndNil(FSer);
  if FBandFmts <> Nil then 
    FreeAndNil(FBandFmts);
end;

procedure TEG_SurfaceChartShared.Assign(AItem: TEG_SurfaceChartShared);
begin
end;

procedure TEG_SurfaceChartShared.CopyTo(AItem: TEG_SurfaceChartShared);
begin
end;

function  TEG_SurfaceChartShared.Create_Wireframe: TCT_Boolean;
begin
  Result := TCT_Boolean.Create(FOwner);
  FWireframe := Result;
end;

function  TEG_SurfaceChartShared.Create_Ser: TCT_SurfaceSer;
begin
  Result := TCT_SurfaceSer.Create(FOwner);
  FSer := Result;
end;

function  TEG_SurfaceChartShared.Create_BandFmts: TCT_BandFmts;
begin
  Result := TCT_BandFmts.Create(FOwner);
  FBandFmts := Result;
end;

{ TCT_BubbleSer }

function  TCT_BubbleSer.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FEG_SerShared.CheckAssigned);
  if FInvertIfNegative <> Nil then 
    Inc(ElemsAssigned,FInvertIfNegative.CheckAssigned);
  if FDPtXpgList <> Nil then 
    Inc(ElemsAssigned,FDPtXpgList.CheckAssigned);
  if FDLbls <> Nil then 
    Inc(ElemsAssigned,FDLbls.CheckAssigned);
  if FTrendlineXpgList <> Nil then 
    Inc(ElemsAssigned,FTrendlineXpgList.CheckAssigned);
  if FErrBarsXpgList <> Nil then 
    Inc(ElemsAssigned,FErrBarsXpgList.CheckAssigned);
  if FXVal <> Nil then 
    Inc(ElemsAssigned,FXVal.CheckAssigned);
  if FYVal <> Nil then 
    Inc(ElemsAssigned,FYVal.CheckAssigned);
  if FBubbleSize <> Nil then 
    Inc(ElemsAssigned,FBubbleSize.CheckAssigned);
  if FBubble3D <> Nil then 
    Inc(ElemsAssigned,FBubble3D.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_BubbleSer.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $00000717: begin
      if FInvertIfNegative = Nil then 
        FInvertIfNegative := TCT_Boolean.Create(FOwner);
      Result := FInvertIfNegative;
    end;
    $000001C5: begin
      if FDPtXpgList = Nil then 
        FDPtXpgList := TCT_DPtXpgList.Create(FOwner);
      Result := FDPtXpgList.Add;
    end;
    $0000028E: begin
      if FDLbls = Nil then 
        FDLbls := TCT_DLbls.Create(FOwner);
      Result := FDLbls;
    end;
    $00000462: begin
      if FTrendlineXpgList = Nil then 
        FTrendlineXpgList := TCT_TrendlineXpgList.Create(FOwner);
      Result := FTrendlineXpgList.Add;
    end;
    $0000036E: begin
      if FErrBarsXpgList = Nil then 
        FErrBarsXpgList := TCT_ErrBarsXpgList.Create(FOwner);
      Result := FErrBarsXpgList.Add;
    end;
    $00000238: begin
      if FXVal = Nil then 
        FXVal := TCT_AxDataSource.Create(FOwner);
      Result := FXVal;
    end;
    $00000239: begin
      if FYVal = Nil then 
        FYVal := TCT_NumDataSource.Create(FOwner);
      Result := FYVal;
    end;
    $000004A4: begin
      if FBubbleSize = Nil then 
        FBubbleSize := TCT_NumDataSource.Create(FOwner);
      Result := FBubbleSize;
    end;
    $00000380: begin
      if FBubble3D = Nil then 
        FBubble3D := TCT_Boolean.Create(FOwner);
      Result := FBubble3D;
    end;
    $00000321: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
    begin
      Result := FEG_SerShared.HandleElement(AReader);
      if Result = Nil then 
        FOwner.Errors.Error(xemUnknownElement,AReader.QName);
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_BubbleSer.Write(AWriter: TXpgWriteXML);
begin
  FEG_SerShared.Write(AWriter);
  if (FInvertIfNegative <> Nil) and FInvertIfNegative.Assigned then 
  begin
    FInvertIfNegative.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:invertIfNegative');
  end;
  if FDPtXpgList <> Nil then 
    FDPtXpgList.Write(AWriter,'c:dPt');
  if (FDLbls <> Nil) and FDLbls.Assigned then 
    if xaElements in FDLbls.Assigneds then
    begin
      AWriter.BeginTag('c:dLbls');
      FDLbls.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:dLbls');
  if FTrendlineXpgList <> Nil then 
    FTrendlineXpgList.Write(AWriter,'c:trendline');
  if FErrBarsXpgList <> Nil then 
    FErrBarsXpgList.Write(AWriter,'c:errBars');
  if (FXVal <> Nil) and FXVal.Assigned then 
    if xaElements in FXVal.Assigneds then
    begin
      AWriter.BeginTag('c:xVal');
      FXVal.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:xVal');
  if (FYVal <> Nil) and FYVal.Assigned then 
    if xaElements in FYVal.Assigneds then
    begin
      AWriter.BeginTag('c:yVal');
      FYVal.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:yVal');
  if (FBubbleSize <> Nil) and FBubbleSize.Assigned then 
    if xaElements in FBubbleSize.Assigneds then
    begin
      AWriter.BeginTag('c:bubbleSize');
      FBubbleSize.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:bubbleSize');
  if (FBubble3D <> Nil) and FBubble3D.Assigned then 
  begin
    FBubble3D.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:bubble3D');
  end;
{$ifdef XLS_WRITE_EXTLST}
  if (FExtLst <> Nil) and FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then
    begin
      AWriter.BeginTag('c:extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:extLst');
{$endif}
end;

constructor TCT_BubbleSer.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 11;
  FAttributeCount := 0;
  FEG_SerShared := TEG_SerShared.Create(FOwner);
end;

destructor TCT_BubbleSer.Destroy;
begin
  FEG_SerShared.Free;
  if FInvertIfNegative <> Nil then 
    FInvertIfNegative.Free;
  if FDPtXpgList <> Nil then 
    FDPtXpgList.Free;
  if FDLbls <> Nil then 
    FDLbls.Free;
  if FTrendlineXpgList <> Nil then 
    FTrendlineXpgList.Free;
  if FErrBarsXpgList <> Nil then 
    FErrBarsXpgList.Free;
  if FXVal <> Nil then 
    FXVal.Free;
  if FYVal <> Nil then 
    FYVal.Free;
  if FBubbleSize <> Nil then 
    FBubbleSize.Free;
  if FBubble3D <> Nil then 
    FBubble3D.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_BubbleSer.Clear;
begin
  FAssigneds := [];
  FEG_SerShared.Clear;
  if FInvertIfNegative <> Nil then 
    FreeAndNil(FInvertIfNegative);
  if FDPtXpgList <> Nil then 
    FreeAndNil(FDPtXpgList);
  if FDLbls <> Nil then 
    FreeAndNil(FDLbls);
  if FTrendlineXpgList <> Nil then 
    FreeAndNil(FTrendlineXpgList);
  if FErrBarsXpgList <> Nil then 
    FreeAndNil(FErrBarsXpgList);
  if FXVal <> Nil then 
    FreeAndNil(FXVal);
  if FYVal <> Nil then 
    FreeAndNil(FYVal);
  if FBubbleSize <> Nil then 
    FreeAndNil(FBubbleSize);
  if FBubble3D <> Nil then 
    FreeAndNil(FBubble3D);
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
end;

procedure TCT_BubbleSer.Assign(AItem: TCT_BubbleSer);
begin
end;

procedure TCT_BubbleSer.CopyTo(AItem: TCT_BubbleSer);
begin
end;

function  TCT_BubbleSer.Create_InvertIfNegative: TCT_Boolean;
begin
  Result := TCT_Boolean.Create(FOwner);
  FInvertIfNegative := Result;
end;

function  TCT_BubbleSer.Create_DPtXpgList: TCT_DPtXpgList;
begin
  Result := TCT_DPtXpgList.Create(FOwner);
  FDPtXpgList := Result;
end;

function  TCT_BubbleSer.Create_DLbls: TCT_DLbls;
begin
  Result := TCT_DLbls.Create(FOwner);
  FDLbls := Result;
end;

function  TCT_BubbleSer.Create_TrendlineXpgList: TCT_TrendlineXpgList;
begin
  Result := TCT_TrendlineXpgList.Create(FOwner);
  FTrendlineXpgList := Result;
end;

function  TCT_BubbleSer.Create_ErrBarsXpgList: TCT_ErrBarsXpgList;
begin
  Result := TCT_ErrBarsXpgList.Create(FOwner);
  FErrBarsXpgList := Result;
end;

function  TCT_BubbleSer.Create_XVal: TCT_AxDataSource;
begin
  Result := TCT_AxDataSource.Create(FOwner);
  FXVal := Result;
end;

function  TCT_BubbleSer.Create_YVal: TCT_NumDataSource;
begin
  Result := TCT_NumDataSource.Create(FOwner);
  FYVal := Result;
end;

function  TCT_BubbleSer.Create_BubbleSize: TCT_NumDataSource;
begin
  Result := TCT_NumDataSource.Create(FOwner);
  FBubbleSize := Result;
end;

function  TCT_BubbleSer.Create_Bubble3D: TCT_Boolean;
begin
  Result := TCT_Boolean.Create(FOwner);
  FBubble3D := Result;
end;

function  TCT_BubbleSer.Create_ExtLst: TCT_ExtensionList;
begin
  Result := TCT_ExtensionList.Create(FOwner);
  FExtLst := Result;
end;

{ TCT_BubbleSerXpgList }

function  TCT_BubbleSerXpgList.GetItems(Index: integer): TCT_BubbleSer;
begin
  Result := TCT_BubbleSer(inherited Items[Index]);
end;

function  TCT_BubbleSerXpgList.Add: TCT_BubbleSer;
begin
  Result := TCT_BubbleSer.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_BubbleSerXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_BubbleSerXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
    if xaElements in Items[i].Assigneds then
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
end;

procedure TCT_BubbleSerXpgList.Assign(AItem: TCT_BubbleSerXpgList);
begin
end;

procedure TCT_BubbleSerXpgList.CopyTo(AItem: TCT_BubbleSerXpgList);
begin
end;

{ TCT_BubbleScale }

function  TCT_BubbleScale.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FVal <> 100 then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_BubbleScale.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_BubbleScale.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FVal <> 100 then 
    AWriter.AddAttribute('val',XmlIntToStr(FVal));
end;

procedure TCT_BubbleScale.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then 
    FVal := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_BubbleScale.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := 100;
end;

destructor TCT_BubbleScale.Destroy;
begin
end;

procedure TCT_BubbleScale.Clear;
begin
  FAssigneds := [];
  FVal := 100;
end;

procedure TCT_BubbleScale.Assign(AItem: TCT_BubbleScale);
begin
end;

procedure TCT_BubbleScale.CopyTo(AItem: TCT_BubbleScale);
begin
end;

{ TCT_SizeRepresents }

function  TCT_SizeRepresents.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FVal <> stsrArea then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_SizeRepresents.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_SizeRepresents.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FVal <> stsrArea then 
    AWriter.AddAttribute('val',StrTST_SizeRepresents[Integer(FVal)]);
end;

procedure TCT_SizeRepresents.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then 
    FVal := TST_SizeRepresents(StrToEnum('stsr' + AAttributes.Values[0]))
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_SizeRepresents.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := stsrArea;
end;

destructor TCT_SizeRepresents.Destroy;
begin
end;

procedure TCT_SizeRepresents.Clear;
begin
  FAssigneds := [];
  FVal := stsrArea;
end;

procedure TCT_SizeRepresents.Assign(AItem: TCT_SizeRepresents);
begin
end;

procedure TCT_SizeRepresents.CopyTo(AItem: TCT_SizeRepresents);
begin
end;

{ TEG_AxShared }

function  TEG_AxShared.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FAxId <> Nil then
    Inc(ElemsAssigned,FAxId.CheckAssigned);
  if FScaling <> Nil then
    Inc(ElemsAssigned,FScaling.CheckAssigned);
  if FDelete <> Nil then
    Inc(ElemsAssigned,FDelete.CheckAssigned);
  if FAxPos <> Nil then
    Inc(ElemsAssigned,FAxPos.CheckAssigned);
  if FMajorGridlines <> Nil then
    Inc(ElemsAssigned);
  if FMinorGridlines <> Nil then
    Inc(ElemsAssigned);
  if FTitle <> Nil then 
    Inc(ElemsAssigned,FTitle.CheckAssigned);
  if FNumFmt <> Nil then 
    Inc(ElemsAssigned,FNumFmt.CheckAssigned);
  if FMajorTickMark <> Nil then 
    Inc(ElemsAssigned,FMajorTickMark.CheckAssigned);
  if FMinorTickMark <> Nil then 
    Inc(ElemsAssigned,FMinorTickMark.CheckAssigned);
  if FTickLblPos <> Nil then 
    Inc(ElemsAssigned,FTickLblPos.CheckAssigned);
  if FSpPr <> Nil then 
    Inc(ElemsAssigned,FSpPr.CheckAssigned);
  if FTxPr <> Nil then 
    Inc(ElemsAssigned,FTxPr.CheckAssigned);
  if FCrossAx <> Nil then 
    Inc(ElemsAssigned,FCrossAx.CheckAssigned);
  if FCrosses <> Nil then 
    Inc(ElemsAssigned,FCrosses.CheckAssigned);
  if FCrossesAt <> Nil then 
    Inc(ElemsAssigned,FCrossesAt.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TEG_AxShared.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $00000223: begin
      if FAxId = Nil then 
        FAxId := TCT_UnsignedInt.Create(FOwner);
      Result := FAxId;
    end;
    $0000037E: begin
      if FScaling = Nil then 
        FScaling := TCT_Scaling.Create(FOwner);
      Result := FScaling;
    end;
    $00000310: begin
      if FDelete = Nil then 
        FDelete := TCT_Boolean.Create(FOwner);
      Result := FDelete;
    end;
    $000002A8: begin
      if FAxPos = Nil then 
        FAxPos := TCT_AxPos.Create(FOwner);
      Result := FAxPos;
    end;
    $00000657: begin
      if FMajorGridlines = Nil then 
        FMajorGridlines := TCT_ChartLines.Create(FOwner);
      Result := FMajorGridlines;
    end;
    $00000663: begin
      if FMinorGridlines = Nil then 
        FMinorGridlines := TCT_ChartLines.Create(FOwner);
      Result := FMinorGridlines;
    end;
    $000002BF: begin
      if FTitle = Nil then 
        FTitle := TCT_Title.Create(FOwner);
      Result := FTitle;
    end;
    $00000314: begin
      if FNumFmt = Nil then 
        FNumFmt := TCT_NumFmt.Create(FOwner);
      Result := FNumFmt;
    end;
    $000005CC: begin
      if FMajorTickMark = Nil then 
        FMajorTickMark := TCT_TickMark.Create(FOwner);
      Result := FMajorTickMark;
    end;
    $000005D8: begin
      if FMinorTickMark = Nil then 
        FMinorTickMark := TCT_TickMark.Create(FOwner);
      Result := FMinorTickMark;
    end;
    $00000494: begin
      if FTickLblPos = Nil then 
        FTickLblPos := TCT_TickLblPos.Create(FOwner);
      Result := FTickLblPos;
    end;
    $00000242: begin
      if FSpPr = Nil then 
        FSpPr := TCT_ShapeProperties.Create(FOwner);
      Result := FSpPr;
    end;
    $0000024B: begin
      if FTxPr = Nil then 
        FTxPr := TCT_TextBody.Create(FOwner);
      Result := FTxPr;
    end;
    $00000380: begin
      if FCrossAx = Nil then 
        FCrossAx := TCT_UnsignedInt.Create(FOwner);
      Result := FCrossAx;
    end;
    $0000039F: begin
      if FCrosses = Nil then 
        FCrosses := TCT_Crosses.Create(FOwner);
      Result := FCrosses;
    end;
    $00000454: begin
      if FCrossesAt = Nil then
        FCrossesAt := TCT_Double.Create(FOwner);
      Result := FCrossesAt;
    end;
    else
    begin
      Result := Nil;
      Exit;
    end
  end;
  if Result <> Self then
    Result.Assigneds := [xaRead];
end;

procedure TEG_AxShared.Write(AWriter: TXpgWriteXML);
begin
  if FAxId.Assigned then
    FAxId.WriteAttributes(AWriter);
  AWriter.SimpleTag('c:axId');

// TODO
//  if FScaling.Assigned then begin
    AWriter.BeginTag('c:scaling');
    FScaling.Write(AWriter);
    AWriter.EndTag;
//  end
//  else
//    AWriter.SimpleTag('c:scaling');

  if (FDelete <> Nil) and FDelete.Assigned then
  begin
    FDelete.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:delete');
  end;

  if FAxPos.Assigned then
    FAxPos.WriteAttributes(AWriter);
  AWriter.SimpleTag('c:axPos');

  if FMajorGridlines <> Nil then begin
    if FMajorGridlines.Assigned and (xaElements in FMajorGridlines.Assigneds) then
    begin
      AWriter.BeginTag('c:majorGridlines');
      FMajorGridlines.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:majorGridlines');
  end;

  if FMinorGridlines <> Nil then begin
    if FMinorGridlines.Assigned and (xaElements in FMinorGridlines.Assigneds) then
    begin
      AWriter.BeginTag('c:minorGridlines');
      FMinorGridlines.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:minorGridlines');
  end;

  if (FTitle <> Nil) and FTitle.Assigned then
    if xaElements in FTitle.Assigneds then
    begin
      AWriter.BeginTag('c:title');
      FTitle.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:title');
  if (FNumFmt <> Nil) and FNumFmt.Assigned then
  begin
    FNumFmt.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:numFmt');
  end;
  if (FMajorTickMark <> Nil) and FMajorTickMark.Assigned then
  begin
    FMajorTickMark.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:majorTickMark');
  end;
  if (FMinorTickMark <> Nil) and FMinorTickMark.Assigned then
  begin
    FMinorTickMark.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:minorTickMark');
  end;
// TODO
//  if (FTickLblPos <> Nil) and FTickLblPos.Assigned then
//  begin
    FTickLblPos.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:tickLblPos');
//  end;
  if (FSpPr <> Nil) and FSpPr.Assigned then
  begin
    FSpPr.WriteAttributes(AWriter);
    if xaElements in FSpPr.Assigneds then
    begin
      AWriter.BeginTag('c:spPr');
      FSpPr.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:spPr');
  end;
  if (FTxPr <> Nil) and FTxPr.Assigned then
    if xaElements in FTxPr.Assigneds then
    begin
      AWriter.BeginTag('c:txPr');
      FTxPr.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:txPr');

  if FCrossAx.Assigned then
    FCrossAx.WriteAttributes(AWriter);
  AWriter.SimpleTag('c:crossAx');

  if (FCrosses <> Nil) and FCrosses.Assigned then
  begin
    FCrosses.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:crosses');
  end;
  if (FCrossesAt <> Nil) and FCrossesAt.Assigned then
  begin
    FCrossesAt.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:crossesAt');
  end;
end;

constructor TEG_AxShared.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 16;
  FAttributeCount := 0;

  FAxId := TCT_UnsignedInt.Create(FOwner);
  FScaling := TCT_Scaling.Create(FOwner);
  FAxPos := TCT_AxPos.Create(FOwner);
  FCrossAx := TCT_UnsignedInt.Create(FOwner);
end;

destructor TEG_AxShared.Destroy;
begin
  FAxId.Free;
  FScaling.Free;
  if FDelete <> Nil then
    FDelete.Free;
  FAxPos.Free;
  if FMajorGridlines <> Nil then
    FMajorGridlines.Free;
  if FMinorGridlines <> Nil then
    FMinorGridlines.Free;
  if FTitle <> Nil then
    FTitle.Free;
  if FNumFmt <> Nil then
    FNumFmt.Free;
  if FMajorTickMark <> Nil then
    FMajorTickMark.Free;
  if FMinorTickMark <> Nil then
    FMinorTickMark.Free;
  if FTickLblPos <> Nil then
    FTickLblPos.Free;
  if FSpPr <> Nil then
    FSpPr.Free;
  if FTxPr <> Nil then
    FTxPr.Free;
  FCrossAx.Free;
  if FCrosses <> Nil then
    FCrosses.Free;
  if FCrossesAt <> Nil then
    FCrossesAt.Free;
end;

procedure TEG_AxShared.Clear;
begin
  FAssigneds := [];
  FAxId.Clear;
  FScaling.Clear;
  if FDelete <> Nil then
    FreeAndNil(FDelete);
  FAxPos.Clear;
  if FMajorGridlines <> Nil then
    FreeAndNil(FMajorGridlines);
  if FMinorGridlines <> Nil then
    FreeAndNil(FMinorGridlines);
  if FTitle <> Nil then
    FreeAndNil(FTitle);
  if FNumFmt <> Nil then
    FreeAndNil(FNumFmt);
  if FMajorTickMark <> Nil then
    FreeAndNil(FMajorTickMark);
  if FMinorTickMark <> Nil then
    FreeAndNil(FMinorTickMark);
  if FTickLblPos <> Nil then
    FreeAndNil(FTickLblPos);
  if FSpPr <> Nil then
    FreeAndNil(FSpPr);
  if FTxPr <> Nil then
    FreeAndNil(FTxPr);
  FCrossAx.Clear;
  if FCrosses <> Nil then
    FreeAndNil(FCrosses);
  if FCrossesAt <> Nil then
    FCrossesAt.Clear;
end;

procedure TEG_AxShared.Assign(AItem: TEG_AxShared);
begin
end;

procedure TEG_AxShared.CopyTo(AItem: TEG_AxShared);
begin
end;

function  TEG_AxShared.Create_Delete: TCT_Boolean;
begin
  Result := TCT_Boolean.Create(FOwner);
  FDelete := Result;
end;

function  TEG_AxShared.Create_MajorGridlines: TCT_ChartLines;
begin
  Result := TCT_ChartLines.Create(FOwner);
  FMajorGridlines := Result;
end;

function  TEG_AxShared.Create_MinorGridlines: TCT_ChartLines;
begin
  Result := TCT_ChartLines.Create(FOwner);
  FMinorGridlines := Result;
end;

function  TEG_AxShared.Create_Title: TCT_Title;
begin
  Result := TCT_Title.Create(FOwner);
  FTitle := Result;
end;

function  TEG_AxShared.Create_NumFmt: TCT_NumFmt;
begin
  Result := TCT_NumFmt.Create(FOwner);
  FNumFmt := Result;
end;

function  TEG_AxShared.Create_MajorTickMark: TCT_TickMark;
begin
  Result := TCT_TickMark.Create(FOwner);
  FMajorTickMark := Result;
end;

function  TEG_AxShared.Create_MinorTickMark: TCT_TickMark;
begin
  Result := TCT_TickMark.Create(FOwner);
  FMinorTickMark := Result;
end;

function  TEG_AxShared.Create_TickLblPos: TCT_TickLblPos;
begin
  Result := TCT_TickLblPos.Create(FOwner);
  FTickLblPos := Result;
end;

function  TEG_AxShared.Create_SpPr: TCT_ShapeProperties;
begin
  Result := TCT_ShapeProperties.Create(FOwner);
  FSpPr := Result;
end;

function  TEG_AxShared.Create_TxPr: TCT_TextBody;
begin
  Result := TCT_TextBody.Create(FOwner);
  FTxPr := Result;
end;

function  TEG_AxShared.Create_Crosses: TCT_Crosses;
begin
  Result := TCT_Crosses.Create(FOwner);
  FCrosses := Result;
end;

function  TEG_AxShared.Create_CrossesAt: TCT_Double;
begin
  Result := TCT_Double.Create(FOwner);
  FCrossesAt := Result;
end;

{ TCT_CrossBetween }

function  TCT_CrossBetween.CheckAssigned: integer;
begin
  FAssigneds := [xaAttributes];
  if FVal <> TST_CrossBetween(XPG_UNKNOWN_ENUM) then
    Result := 1
  else
    Result := 0;
end;

procedure TCT_CrossBetween.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_CrossBetween.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FVal <> TST_CrossBetween(XPG_UNKNOWN_ENUM) then
    AWriter.AddAttribute('val',StrTST_CrossBetween[Integer(FVal)]);
end;

procedure TCT_CrossBetween.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then
    FVal := TST_CrossBetween(StrToEnum('stcb' + AAttributes.Values[0]))
  else
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_CrossBetween.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  Clear;
end;

destructor TCT_CrossBetween.Destroy;
begin
end;

procedure TCT_CrossBetween.Clear;
begin
  FAssigneds := [];
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := TST_CrossBetween(XPG_UNKNOWN_ENUM);
end;

procedure TCT_CrossBetween.Assign(AItem: TCT_CrossBetween);
begin
end;

procedure TCT_CrossBetween.CopyTo(AItem: TCT_CrossBetween);
begin
end;

{ TCT_AxisUnit }

function  TCT_AxisUnit.CheckAssigned: integer;
begin
  FAssigneds := [xaAttributes];
  Result := 1;
end;

procedure TCT_AxisUnit.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_AxisUnit.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('val',XmlFloatToStr(FVal));
end;

procedure TCT_AxisUnit.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then 
    FVal := XmlStrToFloatDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_AxisUnit.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := NaN;
end;

destructor TCT_AxisUnit.Destroy;
begin
end;

procedure TCT_AxisUnit.Clear;
begin
  FAssigneds := [];
  FVal := NaN;
end;

procedure TCT_AxisUnit.Assign(AItem: TCT_AxisUnit);
begin
end;

procedure TCT_AxisUnit.CopyTo(AItem: TCT_AxisUnit);
begin
end;

{ TCT_DispUnits }

function  TCT_DispUnits.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FCustUnit <> Nil then 
    Inc(ElemsAssigned,FCustUnit.CheckAssigned);
  if FBuiltInUnit <> Nil then 
    Inc(ElemsAssigned,FBuiltInUnit.CheckAssigned);
  if FDispUnitsLbl <> Nil then 
    Inc(ElemsAssigned,FDispUnitsLbl.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_DispUnits.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $000003FC: begin
      if FCustUnit = Nil then 
        FCustUnit := TCT_Double.Create(FOwner);
      Result := FCustUnit;
    end;
    $00000514: begin
      if FBuiltInUnit = Nil then 
        FBuiltInUnit := TCT_BuiltInUnit.Create(FOwner);
      Result := FBuiltInUnit;
    end;
    $0000057A: begin
      if FDispUnitsLbl = Nil then 
        FDispUnitsLbl := TCT_DispUnitsLbl.Create(FOwner);
      Result := FDispUnitsLbl;
    end;
    $00000321: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_DispUnits.Write(AWriter: TXpgWriteXML);
begin
  if (FCustUnit <> Nil) and FCustUnit.Assigned then 
  begin
    FCustUnit.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:custUnit');
  end;
// TODO
//  else
//    AWriter.SimpleTag('c:custUnit');
  if (FBuiltInUnit <> Nil) and FBuiltInUnit.Assigned then
  begin
    FBuiltInUnit.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:builtInUnit');
  end
  else 
    AWriter.SimpleTag('c:builtInUnit');
  if (FDispUnitsLbl <> Nil) and FDispUnitsLbl.Assigned then 
    if xaElements in FDispUnitsLbl.Assigneds then
    begin
      AWriter.BeginTag('c:dispUnitsLbl');
      FDispUnitsLbl.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:dispUnitsLbl');
{$ifdef XLS_WRITE_EXTLST}
  if (FExtLst <> Nil) and FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then
    begin
      AWriter.BeginTag('c:extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:extLst');
{$endif}
end;

constructor TCT_DispUnits.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 4;
  FAttributeCount := 0;
end;

destructor TCT_DispUnits.Destroy;
begin
  if FCustUnit <> Nil then 
    FCustUnit.Free;
  if FBuiltInUnit <> Nil then 
    FBuiltInUnit.Free;
  if FDispUnitsLbl <> Nil then 
    FDispUnitsLbl.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_DispUnits.Clear;
begin
  FAssigneds := [];
  if FCustUnit <> Nil then 
    FreeAndNil(FCustUnit);
  if FBuiltInUnit <> Nil then 
    FreeAndNil(FBuiltInUnit);
  if FDispUnitsLbl <> Nil then 
    FreeAndNil(FDispUnitsLbl);
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
end;

procedure TCT_DispUnits.Assign(AItem: TCT_DispUnits);
begin
end;

procedure TCT_DispUnits.CopyTo(AItem: TCT_DispUnits);
begin
end;

function  TCT_DispUnits.Create_CustUnit: TCT_Double;
begin
  Result := TCT_Double.Create(FOwner);
  FCustUnit := Result;
end;

function  TCT_DispUnits.Create_BuiltInUnit: TCT_BuiltInUnit;
begin
  Result := TCT_BuiltInUnit.Create(FOwner);
  FBuiltInUnit := Result;
end;

function  TCT_DispUnits.Create_DispUnitsLbl: TCT_DispUnitsLbl;
begin
  Result := TCT_DispUnitsLbl.Create(FOwner);
  FDispUnitsLbl := Result;
end;

function  TCT_DispUnits.Create_ExtLst: TCT_ExtensionList;
begin
  Result := TCT_ExtensionList.Create(FOwner);
  FExtLst := Result;
end;

{ TCT_LblAlgn }

function  TCT_LblAlgn.CheckAssigned: integer;
begin
  FAssigneds := [xaAttributes];
  if FVal <> TST_LblAlgn(XPG_UNKNOWN_ENUM) then
    Result := 1
  else
    Result := 0;
end;

procedure TCT_LblAlgn.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_LblAlgn.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FVal <> TST_LblAlgn(XPG_UNKNOWN_ENUM) then
    AWriter.AddAttribute('val',StrTST_LblAlgn[Integer(FVal)]);
end;

procedure TCT_LblAlgn.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then
    FVal := TST_LblAlgn(StrToEnum('stla' + AAttributes.Values[0]))
  else
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_LblAlgn.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  Clear;
end;

destructor TCT_LblAlgn.Destroy;
begin
end;

procedure TCT_LblAlgn.Clear;
begin
  FAssigneds := [];
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := TST_LblAlgn(XPG_UNKNOWN_ENUM);
end;

procedure TCT_LblAlgn.Assign(AItem: TCT_LblAlgn);
begin
end;

procedure TCT_LblAlgn.CopyTo(AItem: TCT_LblAlgn);
begin
end;

{ TCT_LblOffset }

function  TCT_LblOffset.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FVal <> 100 then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_LblOffset.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_LblOffset.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FVal <> 100 then 
    AWriter.AddAttribute('val',XmlIntToStr(FVal));
end;

procedure TCT_LblOffset.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then 
    FVal := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_LblOffset.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := 100;
end;

destructor TCT_LblOffset.Destroy;
begin
end;

procedure TCT_LblOffset.Clear;
begin
  FAssigneds := [];
  FVal := 100;
end;

procedure TCT_LblOffset.Assign(AItem: TCT_LblOffset);
begin
end;

procedure TCT_LblOffset.CopyTo(AItem: TCT_LblOffset);
begin
end;

{ TCT_Skip }

function  TCT_Skip.CheckAssigned: integer;
begin
  FAssigneds := [xaAttributes];
  Result := 1;
end;

procedure TCT_Skip.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_Skip.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('val',XmlIntToStr(FVal));
end;

procedure TCT_Skip.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then 
    FVal := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_Skip.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := 2147483632;
end;

destructor TCT_Skip.Destroy;
begin
end;

procedure TCT_Skip.Clear;
begin
  FAssigneds := [];
  FVal := 2147483632;
end;

procedure TCT_Skip.Assign(AItem: TCT_Skip);
begin
end;

procedure TCT_Skip.CopyTo(AItem: TCT_Skip);
begin
end;

{ TCT_TimeUnit }

function  TCT_TimeUnit.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FVal <> sttuDays then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_TimeUnit.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_TimeUnit.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FVal <> sttuDays then 
    AWriter.AddAttribute('val',StrTST_TimeUnit[Integer(FVal)]);
end;

procedure TCT_TimeUnit.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then 
    FVal := TST_TimeUnit(StrToEnum('sttu' + AAttributes.Values[0]))
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_TimeUnit.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := sttuDays;
end;

destructor TCT_TimeUnit.Destroy;
begin
end;

procedure TCT_TimeUnit.Clear;
begin
  FAssigneds := [];
  FVal := sttuDays;
end;

procedure TCT_TimeUnit.Assign(AItem: TCT_TimeUnit);
begin
end;

procedure TCT_TimeUnit.CopyTo(AItem: TCT_TimeUnit);
begin
end;

{ TEG_LegendEntryData }

function  TEG_LegendEntryData.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FTxPr <> Nil then 
    Inc(ElemsAssigned,FTxPr.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TEG_LegendEntryData.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'c:txPr' then 
  begin
    if FTxPr = Nil then 
      FTxPr := TCT_TextBody.Create(FOwner);
    Result := FTxPr;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TEG_LegendEntryData.Write(AWriter: TXpgWriteXML);
begin
  if (FTxPr <> Nil) and FTxPr.Assigned then 
    if xaElements in FTxPr.Assigneds then
    begin
      AWriter.BeginTag('c:txPr');
      FTxPr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:txPr');
end;

constructor TEG_LegendEntryData.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
end;

destructor TEG_LegendEntryData.Destroy;
begin
  if FTxPr <> Nil then 
    FTxPr.Free;
end;

procedure TEG_LegendEntryData.Clear;
begin
  FAssigneds := [];
  if FTxPr <> Nil then 
    FreeAndNil(FTxPr);
end;

procedure TEG_LegendEntryData.Assign(AItem: TEG_LegendEntryData);
begin
end;

procedure TEG_LegendEntryData.CopyTo(AItem: TEG_LegendEntryData);
begin
end;

function  TEG_LegendEntryData.Create_TxPr: TCT_TextBody;
begin
  Result := TCT_TextBody.Create(FOwner);
  FTxPr := Result;
end;

{ TCT_ColorMapping }

function  TCT_ColorMapping.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [xaAttributes];
  if FA_ExtLst <> Nil then
    Inc(ElemsAssigned,FA_ExtLst.CheckAssigned);
  Result := 1;
  if ElemsAssigned > 0 then
    FAssigneds := FAssigneds + [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_ColorMapping.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'a:extLst' then 
  begin
    if FA_ExtLst = Nil then 
      FA_ExtLst := TCT_OfficeArtExtensionList.Create(FOwner);
    Result := FA_ExtLst;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_ColorMapping.Write(AWriter: TXpgWriteXML);
begin
  if (FA_ExtLst <> Nil) and FA_ExtLst.Assigned then 
    if xaElements in FA_ExtLst.Assigneds then
    begin
      AWriter.BeginTag('a:extLst');
      FA_ExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('a:extLst');
end;

procedure TCT_ColorMapping.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if TST_ColorSchemeIndex(XPG_UNKNOWN_ENUM) <> FBg1 then
    AWriter.AddAttribute('bg1',StrTST_ColorSchemeIndex[Integer(FBg1)]);
  if TST_ColorSchemeIndex(XPG_UNKNOWN_ENUM) <> FTx1 then
    AWriter.AddAttribute('tx1',StrTST_ColorSchemeIndex[Integer(FTx1)]);
  if TST_ColorSchemeIndex(XPG_UNKNOWN_ENUM) <> FBg2 then
    AWriter.AddAttribute('bg2',StrTST_ColorSchemeIndex[Integer(FBg2)]);
  if TST_ColorSchemeIndex(XPG_UNKNOWN_ENUM) <> FTx2 then
    AWriter.AddAttribute('tx2',StrTST_ColorSchemeIndex[Integer(FTx2)]);
  if TST_ColorSchemeIndex(XPG_UNKNOWN_ENUM) <> FAccent1 then
    AWriter.AddAttribute('accent1',StrTST_ColorSchemeIndex[Integer(FAccent1)]);
  if TST_ColorSchemeIndex(XPG_UNKNOWN_ENUM) <> FAccent2 then
    AWriter.AddAttribute('accent2',StrTST_ColorSchemeIndex[Integer(FAccent2)]);
  if TST_ColorSchemeIndex(XPG_UNKNOWN_ENUM) <> FAccent3 then
    AWriter.AddAttribute('accent3',StrTST_ColorSchemeIndex[Integer(FAccent3)]);
  if TST_ColorSchemeIndex(XPG_UNKNOWN_ENUM) <> FAccent4 then
    AWriter.AddAttribute('accent4',StrTST_ColorSchemeIndex[Integer(FAccent4)]);
  if TST_ColorSchemeIndex(XPG_UNKNOWN_ENUM) <> FAccent5 then
    AWriter.AddAttribute('accent5',StrTST_ColorSchemeIndex[Integer(FAccent5)]);
  if TST_ColorSchemeIndex(XPG_UNKNOWN_ENUM) <> FAccent6 then
    AWriter.AddAttribute('accent6',StrTST_ColorSchemeIndex[Integer(FAccent6)]);
  if TST_ColorSchemeIndex(XPG_UNKNOWN_ENUM) <> FHlink then
    AWriter.AddAttribute('hlink',StrTST_ColorSchemeIndex[Integer(FHlink)]);
  if TST_ColorSchemeIndex(XPG_UNKNOWN_ENUM) <> FFolHlink then
    AWriter.AddAttribute('folHlink',StrTST_ColorSchemeIndex[Integer(FFolHlink)]);
end;

procedure TCT_ColorMapping.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
    case AAttributes.HashA[i] of
      $000000FA: FBg1 := TST_ColorSchemeIndex(StrToEnum('stcsi' + AAttributes.Values[i]));
      $0000011D: FTx1 := TST_ColorSchemeIndex(StrToEnum('stcsi' + AAttributes.Values[i]));
      $000000FB: FBg2 := TST_ColorSchemeIndex(StrToEnum('stcsi' + AAttributes.Values[i]));
      $0000011E: FTx2 := TST_ColorSchemeIndex(StrToEnum('stcsi' + AAttributes.Values[i]));
      $0000029F: FAccent1 := TST_ColorSchemeIndex(StrToEnum('stcsi' + AAttributes.Values[i]));
      $000002A0: FAccent2 := TST_ColorSchemeIndex(StrToEnum('stcsi' + AAttributes.Values[i]));
      $000002A1: FAccent3 := TST_ColorSchemeIndex(StrToEnum('stcsi' + AAttributes.Values[i]));
      $000002A2: FAccent4 := TST_ColorSchemeIndex(StrToEnum('stcsi' + AAttributes.Values[i]));
      $000002A3: FAccent5 := TST_ColorSchemeIndex(StrToEnum('stcsi' + AAttributes.Values[i]));
      $000002A4: FAccent6 := TST_ColorSchemeIndex(StrToEnum('stcsi' + AAttributes.Values[i]));
      $00000216: FHlink := TST_ColorSchemeIndex(StrToEnum('stcsi' + AAttributes.Values[i]));
      $00000337: FFolHlink := TST_ColorSchemeIndex(StrToEnum('stcsi' + AAttributes.Values[i]));
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_ColorMapping.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 12;
  Clear;
end;

destructor TCT_ColorMapping.Destroy;
begin
  if FA_ExtLst <> Nil then
    FA_ExtLst.Free;
end;

procedure TCT_ColorMapping.Clear;
begin
  FAssigneds := [];
  FBg1 := TST_ColorSchemeIndex(XPG_UNKNOWN_ENUM);
  FTx1 := TST_ColorSchemeIndex(XPG_UNKNOWN_ENUM);
  FBg2 := TST_ColorSchemeIndex(XPG_UNKNOWN_ENUM);
  FTx2 := TST_ColorSchemeIndex(XPG_UNKNOWN_ENUM);
  FAccent1 := TST_ColorSchemeIndex(XPG_UNKNOWN_ENUM);
  FAccent2 := TST_ColorSchemeIndex(XPG_UNKNOWN_ENUM);
  FAccent3 := TST_ColorSchemeIndex(XPG_UNKNOWN_ENUM);
  FAccent4 := TST_ColorSchemeIndex(XPG_UNKNOWN_ENUM);
  FAccent5 := TST_ColorSchemeIndex(XPG_UNKNOWN_ENUM);
  FAccent6 := TST_ColorSchemeIndex(XPG_UNKNOWN_ENUM);
  FHlink := TST_ColorSchemeIndex(XPG_UNKNOWN_ENUM);
  FFolHlink := TST_ColorSchemeIndex(XPG_UNKNOWN_ENUM);
  if FA_ExtLst <> Nil then
    FreeAndNil(FA_ExtLst);
end;

procedure TCT_ColorMapping.Assign(AItem: TCT_ColorMapping);
begin
end;

procedure TCT_ColorMapping.CopyTo(AItem: TCT_ColorMapping);
begin
end;

function  TCT_ColorMapping.Create_A_ExtLst: TCT_OfficeArtExtensionList;
begin
  Result := TCT_OfficeArtExtensionList.Create(FOwner);
  FA_ExtLst := Result;
end;

{ TCT_PivotFmt }

function  TCT_PivotFmt.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FIdx <> Nil then 
    Inc(ElemsAssigned,FIdx.CheckAssigned);
  if FSpPr <> Nil then 
    Inc(ElemsAssigned,FSpPr.CheckAssigned);
  if FTxPr <> Nil then 
    Inc(ElemsAssigned,FTxPr.CheckAssigned);
  if FMarker <> Nil then 
    Inc(ElemsAssigned,FMarker.CheckAssigned);
  if FDLbl <> Nil then 
    Inc(ElemsAssigned,FDLbl.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_PivotFmt.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $000001E2: begin
      if FIdx = Nil then 
        FIdx := TCT_UnsignedInt.Create(FOwner);
      Result := FIdx;
    end;
    $00000242: begin
      if FSpPr = Nil then 
        FSpPr := TCT_ShapeProperties.Create(FOwner);
      Result := FSpPr;
    end;
    $0000024B: begin
      if FTxPr = Nil then 
        FTxPr := TCT_TextBody.Create(FOwner);
      Result := FTxPr;
    end;
    $0000031F: begin
      if FMarker = Nil then 
        FMarker := TCT_Marker.Create(FOwner);
      Result := FMarker;
    end;
    $0000021B: begin
      if FDLbl = Nil then 
        FDLbl := TCT_DLbl.Create(FOwner);
      Result := FDLbl;
    end;
    $00000321: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_PivotFmt.Write(AWriter: TXpgWriteXML);
begin
  if (FIdx <> Nil) and FIdx.Assigned then 
  begin
    FIdx.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:idx');
  end
  else 
    AWriter.SimpleTag('c:idx');
  if (FSpPr <> Nil) and FSpPr.Assigned then 
  begin
    FSpPr.WriteAttributes(AWriter);
    if xaElements in FSpPr.Assigneds then
    begin
      AWriter.BeginTag('c:spPr');
      FSpPr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:spPr');
  end;
  if (FTxPr <> Nil) and FTxPr.Assigned then 
    if xaElements in FTxPr.Assigneds then
    begin
      AWriter.BeginTag('c:txPr');
      FTxPr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:txPr');
  if (FMarker <> Nil) and FMarker.Assigned then 
    if xaElements in FMarker.Assigneds then
    begin
      AWriter.BeginTag('c:marker');
      FMarker.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:marker');
  if (FDLbl <> Nil) and FDLbl.Assigned then 
    if xaElements in FDLbl.Assigneds then
    begin
      AWriter.BeginTag('c:dLbl');
      FDLbl.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:dLbl');
{$ifdef XLS_WRITE_EXTLST}
  if (FExtLst <> Nil) and FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then
    begin
      AWriter.BeginTag('c:extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:extLst');
{$endif}
end;

constructor TCT_PivotFmt.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 6;
  FAttributeCount := 0;
end;

destructor TCT_PivotFmt.Destroy;
begin
  if FIdx <> Nil then 
    FIdx.Free;
  if FSpPr <> Nil then 
    FSpPr.Free;
  if FTxPr <> Nil then 
    FTxPr.Free;
  if FMarker <> Nil then 
    FMarker.Free;
  if FDLbl <> Nil then 
    FDLbl.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_PivotFmt.Clear;
begin
  FAssigneds := [];
  if FIdx <> Nil then 
    FreeAndNil(FIdx);
  if FSpPr <> Nil then 
    FreeAndNil(FSpPr);
  if FTxPr <> Nil then 
    FreeAndNil(FTxPr);
  if FMarker <> Nil then 
    FreeAndNil(FMarker);
  if FDLbl <> Nil then 
    FreeAndNil(FDLbl);
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
end;

procedure TCT_PivotFmt.Assign(AItem: TCT_PivotFmt);
begin
end;

procedure TCT_PivotFmt.CopyTo(AItem: TCT_PivotFmt);
begin
end;

function  TCT_PivotFmt.Create_Idx: TCT_UnsignedInt;
begin
  Result := TCT_UnsignedInt.Create(FOwner);
  FIdx := Result;
end;

function  TCT_PivotFmt.Create_SpPr: TCT_ShapeProperties;
begin
  Result := TCT_ShapeProperties.Create(FOwner);
  FSpPr := Result;
end;

function  TCT_PivotFmt.Create_TxPr: TCT_TextBody;
begin
  Result := TCT_TextBody.Create(FOwner);
  FTxPr := Result;
end;

function  TCT_PivotFmt.Create_Marker: TCT_Marker;
begin
  Result := TCT_Marker.Create(FOwner);
  FMarker := Result;
end;

function  TCT_PivotFmt.Create_DLbl: TCT_DLbl;
begin
  Result := TCT_DLbl.Create(FOwner);
  FDLbl := Result;
end;

function  TCT_PivotFmt.Create_ExtLst: TCT_ExtensionList;
begin
  Result := TCT_ExtensionList.Create(FOwner);
  FExtLst := Result;
end;

{ TCT_PivotFmtXpgList }

function  TCT_PivotFmtXpgList.GetItems(Index: integer): TCT_PivotFmt;
begin
  Result := TCT_PivotFmt(inherited Items[Index]);
end;

function  TCT_PivotFmtXpgList.Add: TCT_PivotFmt;
begin
  Result := TCT_PivotFmt.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_PivotFmtXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_PivotFmtXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
    if xaElements in Items[i].Assigneds then
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
end;

procedure TCT_PivotFmtXpgList.Assign(AItem: TCT_PivotFmtXpgList);
begin
end;

procedure TCT_PivotFmtXpgList.CopyTo(AItem: TCT_PivotFmtXpgList);
begin
end;

{ TCT_RotX }

function  TCT_RotX.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FVal <> 0 then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_RotX.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_RotX.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FVal <> 0 then 
    AWriter.AddAttribute('val',XmlIntToStr(FVal));
end;

procedure TCT_RotX.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then 
    FVal := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_RotX.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := 0;
end;

destructor TCT_RotX.Destroy;
begin
end;

procedure TCT_RotX.Clear;
begin
  FAssigneds := [];
  FVal := 0;
end;

procedure TCT_RotX.Assign(AItem: TCT_RotX);
begin
end;

procedure TCT_RotX.CopyTo(AItem: TCT_RotX);
begin
end;

{ TCT_HPercent }

function  TCT_HPercent.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FVal <> 100 then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_HPercent.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_HPercent.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FVal <> 100 then 
    AWriter.AddAttribute('val',XmlIntToStr(FVal));
end;

procedure TCT_HPercent.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then 
    FVal := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_HPercent.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := 100;
end;

destructor TCT_HPercent.Destroy;
begin
end;

procedure TCT_HPercent.Clear;
begin
  FAssigneds := [];
  FVal := 100;
end;

procedure TCT_HPercent.Assign(AItem: TCT_HPercent);
begin
end;

procedure TCT_HPercent.CopyTo(AItem: TCT_HPercent);
begin
end;

{ TCT_RotY }

function  TCT_RotY.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FVal <> 0 then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_RotY.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_RotY.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FVal <> 0 then 
    AWriter.AddAttribute('val',XmlIntToStr(FVal));
end;

procedure TCT_RotY.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then 
    FVal := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_RotY.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := 0;
end;

destructor TCT_RotY.Destroy;
begin
end;

procedure TCT_RotY.Clear;
begin
  FAssigneds := [];
  FVal := 0;
end;

procedure TCT_RotY.Assign(AItem: TCT_RotY);
begin
end;

procedure TCT_RotY.CopyTo(AItem: TCT_RotY);
begin
end;

{ TCT_DepthPercent }

function  TCT_DepthPercent.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FVal <> 100 then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_DepthPercent.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_DepthPercent.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FVal <> 100 then 
    AWriter.AddAttribute('val',XmlIntToStr(FVal));
end;

procedure TCT_DepthPercent.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then 
    FVal := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_DepthPercent.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := 100;
end;

destructor TCT_DepthPercent.Destroy;
begin
end;

procedure TCT_DepthPercent.Clear;
begin
  FAssigneds := [];
  FVal := 100;
end;

procedure TCT_DepthPercent.Assign(AItem: TCT_DepthPercent);
begin
end;

procedure TCT_DepthPercent.CopyTo(AItem: TCT_DepthPercent);
begin
end;

{ TCT_Perspective }

function  TCT_Perspective.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FVal <> 30 then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_Perspective.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_Perspective.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FVal <> 30 then 
    AWriter.AddAttribute('val',XmlIntToStr(FVal));
end;

procedure TCT_Perspective.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then 
    FVal := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_Perspective.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := 30;
end;

destructor TCT_Perspective.Destroy;
begin
end;

procedure TCT_Perspective.Clear;
begin
  FAssigneds := [];
  FVal := 30;
end;

procedure TCT_Perspective.Assign(AItem: TCT_Perspective);
begin
end;

procedure TCT_Perspective.CopyTo(AItem: TCT_Perspective);
begin
end;

{ TCT_AreaChart }

function  TCT_AreaChart.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FEG_AreaChartShared.CheckAssigned);
  if FAxIdXpgList <> Nil then 
    Inc(ElemsAssigned,FAxIdXpgList.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_AreaChart.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $00000223: begin
      if FAxIdXpgList = Nil then 
        FAxIdXpgList := TCT_UnsignedIntXpgList.Create(FOwner);
      Result := FAxIdXpgList.Add;
    end;
    $00000321: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
    begin
      Result := FEG_AreaChartShared.HandleElement(AReader);
      if Result = Nil then 
        FOwner.Errors.Error(xemUnknownElement,AReader.QName);
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_AreaChart.Write(AWriter: TXpgWriteXML);
begin
  FEG_AreaChartShared.Write(AWriter);
  if FAxIdXpgList <> Nil then 
    FAxIdXpgList.Write(AWriter,'c:axId');
{$ifdef XLS_WRITE_EXTLST}
  if (FExtLst <> Nil) and FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then
    begin
      AWriter.BeginTag('c:extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:extLst');
{$endif}
end;

constructor TCT_AreaChart.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 3;
  FAttributeCount := 0;
  FEG_AreaChartShared := TEG_AreaChartShared.Create(FOwner);
end;

destructor TCT_AreaChart.Destroy;
begin
  FEG_AreaChartShared.Free;
  if FAxIdXpgList <> Nil then 
    FAxIdXpgList.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_AreaChart.Clear;
begin
  FAssigneds := [];
  FEG_AreaChartShared.Clear;
  if FAxIdXpgList <> Nil then 
    FreeAndNil(FAxIdXpgList);
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
end;

procedure TCT_AreaChart.Assign(AItem: TCT_AreaChart);
begin
end;

procedure TCT_AreaChart.CopyTo(AItem: TCT_AreaChart);
begin
end;

function  TCT_AreaChart.Create_AxIdXpgList: TCT_UnsignedIntXpgList;
begin
  Result := TCT_UnsignedIntXpgList.Create(FOwner);
  FAxIdXpgList := Result;
end;

function  TCT_AreaChart.Create_ExtLst: TCT_ExtensionList;
begin
  Result := TCT_ExtensionList.Create(FOwner);
  FExtLst := Result;
end;

{ TCT_AreaChartXpgList }

function  TCT_AreaChartXpgList.GetItems(Index: integer): TCT_AreaChart;
begin
  Result := TCT_AreaChart(inherited Items[Index]);
end;

function  TCT_AreaChartXpgList.Add: TCT_AreaChart;
begin
  Result := TCT_AreaChart.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_AreaChartXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_AreaChartXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
    if xaElements in Items[i].Assigneds then
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
end;

procedure TCT_AreaChartXpgList.Assign(AItem: TCT_AreaChartXpgList);
begin
end;

procedure TCT_AreaChartXpgList.CopyTo(AItem: TCT_AreaChartXpgList);
begin
end;

{ TCT_Area3DChart }

function  TCT_Area3DChart.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FEG_AreaChartShared.CheckAssigned);
  if FGapDepth <> Nil then 
    Inc(ElemsAssigned,FGapDepth.CheckAssigned);
  if FAxIdXpgList <> Nil then 
    Inc(ElemsAssigned,FAxIdXpgList.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_Area3DChart.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $000003CA: begin
      if FGapDepth = Nil then 
        FGapDepth := TCT_GapAmount.Create(FOwner);
      Result := FGapDepth;
    end;
    $00000223: begin
      if FAxIdXpgList = Nil then 
        FAxIdXpgList := TCT_UnsignedIntXpgList.Create(FOwner);
      Result := FAxIdXpgList.Add;
    end;
    $00000321: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
    begin
      Result := FEG_AreaChartShared.HandleElement(AReader);
      if Result = Nil then 
        FOwner.Errors.Error(xemUnknownElement,AReader.QName);
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_Area3DChart.Write(AWriter: TXpgWriteXML);
begin
  FEG_AreaChartShared.Write(AWriter);
  if (FGapDepth <> Nil) and FGapDepth.Assigned then 
  begin
    FGapDepth.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:gapDepth');
  end;
  if FAxIdXpgList <> Nil then 
    FAxIdXpgList.Write(AWriter,'c:axId');
{$ifdef XLS_WRITE_EXTLST}
  if (FExtLst <> Nil) and FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then
    begin
      AWriter.BeginTag('c:extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:extLst');
{$endif}
end;

constructor TCT_Area3DChart.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 4;
  FAttributeCount := 0;
  FEG_AreaChartShared := TEG_AreaChartShared.Create(FOwner);
end;

destructor TCT_Area3DChart.Destroy;
begin
  FEG_AreaChartShared.Free;
  if FGapDepth <> Nil then 
    FGapDepth.Free;
  if FAxIdXpgList <> Nil then 
    FAxIdXpgList.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_Area3DChart.Clear;
begin
  FAssigneds := [];
  FEG_AreaChartShared.Clear;
  if FGapDepth <> Nil then 
    FreeAndNil(FGapDepth);
  if FAxIdXpgList <> Nil then 
    FreeAndNil(FAxIdXpgList);
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
end;

procedure TCT_Area3DChart.Assign(AItem: TCT_Area3DChart);
begin
end;

procedure TCT_Area3DChart.CopyTo(AItem: TCT_Area3DChart);
begin
end;

function  TCT_Area3DChart.Create_GapDepth: TCT_GapAmount;
begin
  Result := TCT_GapAmount.Create(FOwner);
  FGapDepth := Result;
end;

function  TCT_Area3DChart.Create_AxIdXpgList: TCT_UnsignedIntXpgList;
begin
  Result := TCT_UnsignedIntXpgList.Create(FOwner);
  FAxIdXpgList := Result;
end;

function  TCT_Area3DChart.Create_ExtLst: TCT_ExtensionList;
begin
  Result := TCT_ExtensionList.Create(FOwner);
  FExtLst := Result;
end;

{ TCT_Area3DChartXpgList }

function  TCT_Area3DChartXpgList.GetItems(Index: integer): TCT_Area3DChart;
begin
  Result := TCT_Area3DChart(inherited Items[Index]);
end;

function  TCT_Area3DChartXpgList.Add: TCT_Area3DChart;
begin
  Result := TCT_Area3DChart.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_Area3DChartXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_Area3DChartXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
    if xaElements in Items[i].Assigneds then
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
end;

procedure TCT_Area3DChartXpgList.Assign(AItem: TCT_Area3DChartXpgList);
begin
end;

procedure TCT_Area3DChartXpgList.CopyTo(AItem: TCT_Area3DChartXpgList);
begin
end;

{ TCT_LineChart }

function  TCT_LineChart.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FEG_LineChartShared.CheckAssigned);
  if FHiLowLines <> Nil then 
    Inc(ElemsAssigned,FHiLowLines.CheckAssigned);
  if FUpDownBars <> Nil then 
    Inc(ElemsAssigned,FUpDownBars.CheckAssigned);
  if FMarker <> Nil then 
    Inc(ElemsAssigned,FMarker.CheckAssigned);
  if FSmooth <> Nil then 
    Inc(ElemsAssigned,FSmooth.CheckAssigned);
  if FAxIdXpgList <> Nil then 
    Inc(ElemsAssigned,FAxIdXpgList.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_LineChart.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $0000049B: begin
      if FHiLowLines = Nil then 
        FHiLowLines := TCT_ChartLines.Create(FOwner);
      Result := FHiLowLines;
    end;
    $000004A2: begin
      if FUpDownBars = Nil then 
        FUpDownBars := TCT_UpDownBars.Create(FOwner);
      Result := FUpDownBars;
    end;
    $0000031F: begin
      if FMarker = Nil then 
        FMarker := TCT_Boolean.Create(FOwner);
      Result := FMarker;
    end;
    $00000337: begin
      if FSmooth = Nil then 
        FSmooth := TCT_Boolean.Create(FOwner);
      Result := FSmooth;
    end;
    $00000223: begin
      if FAxIdXpgList = Nil then 
        FAxIdXpgList := TCT_UnsignedIntXpgList.Create(FOwner);
      Result := FAxIdXpgList.Add;
    end;
    $00000321: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
    begin
      Result := FEG_LineChartShared.HandleElement(AReader);
      if Result = Nil then 
        FOwner.Errors.Error(xemUnknownElement,AReader.QName);
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_LineChart.Write(AWriter: TXpgWriteXML);
begin
  FEG_LineChartShared.Write(AWriter);
  if (FHiLowLines <> Nil) and FHiLowLines.Assigned then 
    if xaElements in FHiLowLines.Assigneds then
    begin
      AWriter.BeginTag('c:hiLowLines');
      FHiLowLines.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:hiLowLines');
  if (FUpDownBars <> Nil) and FUpDownBars.Assigned then 
    if xaElements in FUpDownBars.Assigneds then
    begin
      AWriter.BeginTag('c:upDownBars');
      FUpDownBars.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:upDownBars');
  if (FMarker <> Nil) and FMarker.Assigned then 
  begin
    FMarker.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:marker');
  end;
  if (FSmooth <> Nil) and FSmooth.Assigned then 
  begin
    FSmooth.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:smooth');
  end;
  if FAxIdXpgList <> Nil then 
    FAxIdXpgList.Write(AWriter,'c:axId');
{$ifdef XLS_WRITE_EXTLST}
  if (FExtLst <> Nil) and FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then
    begin
      AWriter.BeginTag('c:extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:extLst');
{$endif}
end;

constructor TCT_LineChart.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 7;
  FAttributeCount := 0;
  FEG_LineChartShared := TEG_LineChartShared.Create(FOwner);
end;

destructor TCT_LineChart.Destroy;
begin
  FEG_LineChartShared.Free;
  if FHiLowLines <> Nil then 
    FHiLowLines.Free;
  if FUpDownBars <> Nil then 
    FUpDownBars.Free;
  if FMarker <> Nil then 
    FMarker.Free;
  if FSmooth <> Nil then 
    FSmooth.Free;
  if FAxIdXpgList <> Nil then 
    FAxIdXpgList.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_LineChart.Clear;
begin
  FAssigneds := [];
  FEG_LineChartShared.Clear;
  if FHiLowLines <> Nil then 
    FreeAndNil(FHiLowLines);
  if FUpDownBars <> Nil then 
    FreeAndNil(FUpDownBars);
  if FMarker <> Nil then 
    FreeAndNil(FMarker);
  if FSmooth <> Nil then 
    FreeAndNil(FSmooth);
  if FAxIdXpgList <> Nil then 
    FreeAndNil(FAxIdXpgList);
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
end;

procedure TCT_LineChart.Assign(AItem: TCT_LineChart);
begin
end;

procedure TCT_LineChart.CopyTo(AItem: TCT_LineChart);
begin
end;

function  TCT_LineChart.Create_HiLowLines: TCT_ChartLines;
begin
  Result := TCT_ChartLines.Create(FOwner);
  FHiLowLines := Result;
end;

function  TCT_LineChart.Create_UpDownBars: TCT_UpDownBars;
begin
  Result := TCT_UpDownBars.Create(FOwner);
  FUpDownBars := Result;
end;

function  TCT_LineChart.Create_Marker: TCT_Boolean;
begin
  Result := TCT_Boolean.Create(FOwner);
  FMarker := Result;
end;

function  TCT_LineChart.Create_Smooth: TCT_Boolean;
begin
  Result := TCT_Boolean.Create(FOwner);
  FSmooth := Result;
end;

function  TCT_LineChart.Create_AxIdXpgList: TCT_UnsignedIntXpgList;
begin
  Result := TCT_UnsignedIntXpgList.Create(FOwner);
  FAxIdXpgList := Result;
end;

function  TCT_LineChart.Create_ExtLst: TCT_ExtensionList;
begin
  Result := TCT_ExtensionList.Create(FOwner);
  FExtLst := Result;
end;

{ TCT_LineChartXpgList }

function  TCT_LineChartXpgList.GetItems(Index: integer): TCT_LineChart;
begin
  Result := TCT_LineChart(inherited Items[Index]);
end;

function  TCT_LineChartXpgList.Add: TCT_LineChart;
begin
  Result := TCT_LineChart.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_LineChartXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_LineChartXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
    if xaElements in Items[i].Assigneds then
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
end;

procedure TCT_LineChartXpgList.Assign(AItem: TCT_LineChartXpgList);
begin
end;

procedure TCT_LineChartXpgList.CopyTo(AItem: TCT_LineChartXpgList);
begin
end;

{ TCT_Line3DChart }

function  TCT_Line3DChart.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FEG_LineChartShared.CheckAssigned);
  if FGapDepth <> Nil then 
    Inc(ElemsAssigned,FGapDepth.CheckAssigned);
  if FAxIdXpgList <> Nil then 
    Inc(ElemsAssigned,FAxIdXpgList.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_Line3DChart.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $000003CA: begin
      if FGapDepth = Nil then 
        FGapDepth := TCT_GapAmount.Create(FOwner);
      Result := FGapDepth;
    end;
    $00000223: begin
      if FAxIdXpgList = Nil then 
        FAxIdXpgList := TCT_UnsignedIntXpgList.Create(FOwner);
      Result := FAxIdXpgList.Add;
    end;
    $00000321: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
    begin
      Result := FEG_LineChartShared.HandleElement(AReader);
      if Result = Nil then 
        FOwner.Errors.Error(xemUnknownElement,AReader.QName);
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_Line3DChart.Write(AWriter: TXpgWriteXML);
begin
  FEG_LineChartShared.Write(AWriter);
  if (FGapDepth <> Nil) and FGapDepth.Assigned then 
  begin
    FGapDepth.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:gapDepth');
  end;
  if FAxIdXpgList <> Nil then 
    FAxIdXpgList.Write(AWriter,'c:axId');
{$ifdef XLS_WRITE_EXTLST}
  if (FExtLst <> Nil) and FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then
    begin
      AWriter.BeginTag('c:extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:extLst');
{$endif}
end;

constructor TCT_Line3DChart.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 4;
  FAttributeCount := 0;
  FEG_LineChartShared := TEG_LineChartShared.Create(FOwner);
end;

destructor TCT_Line3DChart.Destroy;
begin
  FEG_LineChartShared.Free;
  if FGapDepth <> Nil then 
    FGapDepth.Free;
  if FAxIdXpgList <> Nil then 
    FAxIdXpgList.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_Line3DChart.Clear;
begin
  FAssigneds := [];
  FEG_LineChartShared.Clear;
  if FGapDepth <> Nil then 
    FreeAndNil(FGapDepth);
  if FAxIdXpgList <> Nil then 
    FreeAndNil(FAxIdXpgList);
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
end;

procedure TCT_Line3DChart.Assign(AItem: TCT_Line3DChart);
begin
end;

procedure TCT_Line3DChart.CopyTo(AItem: TCT_Line3DChart);
begin
end;

function  TCT_Line3DChart.Create_GapDepth: TCT_GapAmount;
begin
  Result := TCT_GapAmount.Create(FOwner);
  FGapDepth := Result;
end;

function  TCT_Line3DChart.Create_AxIdXpgList: TCT_UnsignedIntXpgList;
begin
  Result := TCT_UnsignedIntXpgList.Create(FOwner);
  FAxIdXpgList := Result;
end;

function  TCT_Line3DChart.Create_ExtLst: TCT_ExtensionList;
begin
  Result := TCT_ExtensionList.Create(FOwner);
  FExtLst := Result;
end;

{ TCT_Line3DChartXpgList }

function  TCT_Line3DChartXpgList.GetItems(Index: integer): TCT_Line3DChart;
begin
  Result := TCT_Line3DChart(inherited Items[Index]);
end;

function  TCT_Line3DChartXpgList.Add: TCT_Line3DChart;
begin
  Result := TCT_Line3DChart.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_Line3DChartXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_Line3DChartXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
    if xaElements in Items[i].Assigneds then
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
end;

procedure TCT_Line3DChartXpgList.Assign(AItem: TCT_Line3DChartXpgList);
begin
end;

procedure TCT_Line3DChartXpgList.CopyTo(AItem: TCT_Line3DChartXpgList);
begin
end;

{ TCT_StockChart }

function  TCT_StockChart.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FSerXpgList <> Nil then 
    Inc(ElemsAssigned,FSerXpgList.CheckAssigned);
  if FDLbls <> Nil then 
    Inc(ElemsAssigned,FDLbls.CheckAssigned);
  if FDropLines <> Nil then 
    Inc(ElemsAssigned,FDropLines.CheckAssigned);
  if FHiLowLines <> Nil then 
    Inc(ElemsAssigned,FHiLowLines.CheckAssigned);
  if FUpDownBars <> Nil then 
    Inc(ElemsAssigned,FUpDownBars.CheckAssigned);
  if FAxIdXpgList <> Nil then 
    Inc(ElemsAssigned,FAxIdXpgList.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_StockChart.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $000001E7: begin
      if FSerXpgList = Nil then 
        FSerXpgList := TCT_LineSerXpgList.Create(FOwner);
      Result := FSerXpgList.Add;
    end;
    $0000028E: begin
      if FDLbls = Nil then 
        FDLbls := TCT_DLbls.Create(FOwner);
      Result := FDLbls;
    end;
    $0000044D: begin
      if FDropLines = Nil then 
        FDropLines := TCT_ChartLines.Create(FOwner);
      Result := FDropLines;
    end;
    $0000049B: begin
      if FHiLowLines = Nil then 
        FHiLowLines := TCT_ChartLines.Create(FOwner);
      Result := FHiLowLines;
    end;
    $000004A2: begin
      if FUpDownBars = Nil then 
        FUpDownBars := TCT_UpDownBars.Create(FOwner);
      Result := FUpDownBars;
    end;
    $00000223: begin
      if FAxIdXpgList = Nil then 
        FAxIdXpgList := TCT_UnsignedIntXpgList.Create(FOwner);
      Result := FAxIdXpgList.Add;
    end;
    $00000321: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_StockChart.Write(AWriter: TXpgWriteXML);
begin
  if FSerXpgList <> Nil then 
    FSerXpgList.Write(AWriter,'c:ser');
  if (FDLbls <> Nil) and FDLbls.Assigned then 
    if xaElements in FDLbls.Assigneds then
    begin
      AWriter.BeginTag('c:dLbls');
      FDLbls.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:dLbls');
  if (FDropLines <> Nil) and FDropLines.Assigned then 
    if xaElements in FDropLines.Assigneds then
    begin
      AWriter.BeginTag('c:dropLines');
      FDropLines.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:dropLines');
  if (FHiLowLines <> Nil) and FHiLowLines.Assigned then 
    if xaElements in FHiLowLines.Assigneds then
    begin
      AWriter.BeginTag('c:hiLowLines');
      FHiLowLines.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:hiLowLines');
  if (FUpDownBars <> Nil) and FUpDownBars.Assigned then 
    if xaElements in FUpDownBars.Assigneds then
    begin
      AWriter.BeginTag('c:upDownBars');
      FUpDownBars.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:upDownBars');
  if FAxIdXpgList <> Nil then 
    FAxIdXpgList.Write(AWriter,'c:axId');
{$ifdef XLS_WRITE_EXTLST}
  if (FExtLst <> Nil) and FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then
    begin
      AWriter.BeginTag('c:extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:extLst');
{$endif}
end;

constructor TCT_StockChart.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 7;
  FAttributeCount := 0;
end;

destructor TCT_StockChart.Destroy;
begin
  if FSerXpgList <> Nil then 
    FSerXpgList.Free;
  if FDLbls <> Nil then 
    FDLbls.Free;
  if FDropLines <> Nil then 
    FDropLines.Free;
  if FHiLowLines <> Nil then 
    FHiLowLines.Free;
  if FUpDownBars <> Nil then 
    FUpDownBars.Free;
  if FAxIdXpgList <> Nil then 
    FAxIdXpgList.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_StockChart.Clear;
begin
  FAssigneds := [];
  if FSerXpgList <> Nil then 
    FreeAndNil(FSerXpgList);
  if FDLbls <> Nil then 
    FreeAndNil(FDLbls);
  if FDropLines <> Nil then 
    FreeAndNil(FDropLines);
  if FHiLowLines <> Nil then 
    FreeAndNil(FHiLowLines);
  if FUpDownBars <> Nil then 
    FreeAndNil(FUpDownBars);
  if FAxIdXpgList <> Nil then 
    FreeAndNil(FAxIdXpgList);
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
end;

procedure TCT_StockChart.Assign(AItem: TCT_StockChart);
begin
end;

procedure TCT_StockChart.CopyTo(AItem: TCT_StockChart);
begin
end;

function  TCT_StockChart.Create_SerXpgList: TCT_LineSerXpgList;
begin
  Result := TCT_LineSerXpgList.Create(FOwner);
  FSerXpgList := Result;
end;

function  TCT_StockChart.Create_DLbls: TCT_DLbls;
begin
  Result := TCT_DLbls.Create(FOwner);
  FDLbls := Result;
end;

function  TCT_StockChart.Create_DropLines: TCT_ChartLines;
begin
  Result := TCT_ChartLines.Create(FOwner);
  FDropLines := Result;
end;

function  TCT_StockChart.Create_HiLowLines: TCT_ChartLines;
begin
  Result := TCT_ChartLines.Create(FOwner);
  FHiLowLines := Result;
end;

function  TCT_StockChart.Create_UpDownBars: TCT_UpDownBars;
begin
  Result := TCT_UpDownBars.Create(FOwner);
  FUpDownBars := Result;
end;

function  TCT_StockChart.Create_AxIdXpgList: TCT_UnsignedIntXpgList;
begin
  Result := TCT_UnsignedIntXpgList.Create(FOwner);
  FAxIdXpgList := Result;
end;

function  TCT_StockChart.Create_ExtLst: TCT_ExtensionList;
begin
  Result := TCT_ExtensionList.Create(FOwner);
  FExtLst := Result;
end;

{ TCT_StockChartXpgList }

function  TCT_StockChartXpgList.GetItems(Index: integer): TCT_StockChart;
begin
  Result := TCT_StockChart(inherited Items[Index]);
end;

function  TCT_StockChartXpgList.Add: TCT_StockChart;
begin
  Result := TCT_StockChart.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_StockChartXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_StockChartXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
    if xaElements in Items[i].Assigneds then
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
end;

procedure TCT_StockChartXpgList.Assign(AItem: TCT_StockChartXpgList);
begin
end;

procedure TCT_StockChartXpgList.CopyTo(AItem: TCT_StockChartXpgList);
begin
end;

{ TCT_RadarChart }

function  TCT_RadarChart.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FRadarStyle <> Nil then 
    Inc(ElemsAssigned,FRadarStyle.CheckAssigned);
  if FVaryColors <> Nil then 
    Inc(ElemsAssigned,FVaryColors.CheckAssigned);
  if FSerXpgList <> Nil then 
    Inc(ElemsAssigned,FSerXpgList.CheckAssigned);
  if FDLbls <> Nil then 
    Inc(ElemsAssigned,FDLbls.CheckAssigned);
  if FAxIdXpgList <> Nil then 
    Inc(ElemsAssigned,FAxIdXpgList.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_RadarChart.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $000004B8: begin
      if FRadarStyle = Nil then 
        FRadarStyle := TCT_RadarStyle.Create(FOwner);
      Result := FRadarStyle;
    end;
    $000004D1: begin
      if FVaryColors = Nil then 
        FVaryColors := TCT_Boolean.Create(FOwner);
      Result := FVaryColors;
    end;
    $000001E7: begin
      if FSerXpgList = Nil then 
        FSerXpgList := TCT_RadarSerXpgList.Create(FOwner);
      Result := FSerXpgList.Add;
    end;
    $0000028E: begin
      if FDLbls = Nil then 
        FDLbls := TCT_DLbls.Create(FOwner);
      Result := FDLbls;
    end;
    $00000223: begin
      if FAxIdXpgList = Nil then 
        FAxIdXpgList := TCT_UnsignedIntXpgList.Create(FOwner);
      Result := FAxIdXpgList.Add;
    end;
    $00000321: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_RadarChart.Write(AWriter: TXpgWriteXML);
begin
  if (FRadarStyle <> Nil) and FRadarStyle.Assigned then 
  begin
    FRadarStyle.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:radarStyle');
  end;
  if (FVaryColors <> Nil) and FVaryColors.Assigned then 
  begin
    FVaryColors.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:varyColors');
  end;
  if FSerXpgList <> Nil then 
    FSerXpgList.Write(AWriter,'c:ser');
  if (FDLbls <> Nil) and FDLbls.Assigned then 
    if xaElements in FDLbls.Assigneds then
    begin
      AWriter.BeginTag('c:dLbls');
      FDLbls.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:dLbls');
  if FAxIdXpgList <> Nil then 
    FAxIdXpgList.Write(AWriter,'c:axId');
{$ifdef XLS_WRITE_EXTLST}
  if (FExtLst <> Nil) and FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then
    begin
      AWriter.BeginTag('c:extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:extLst');
{$endif}
end;

constructor TCT_RadarChart.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 6;
  FAttributeCount := 0;
end;

destructor TCT_RadarChart.Destroy;
begin
  if FRadarStyle <> Nil then 
    FRadarStyle.Free;
  if FVaryColors <> Nil then 
    FVaryColors.Free;
  if FSerXpgList <> Nil then 
    FSerXpgList.Free;
  if FDLbls <> Nil then 
    FDLbls.Free;
  if FAxIdXpgList <> Nil then 
    FAxIdXpgList.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_RadarChart.Clear;
begin
  FAssigneds := [];
  if FRadarStyle <> Nil then 
    FreeAndNil(FRadarStyle);
  if FVaryColors <> Nil then 
    FreeAndNil(FVaryColors);
  if FSerXpgList <> Nil then 
    FreeAndNil(FSerXpgList);
  if FDLbls <> Nil then 
    FreeAndNil(FDLbls);
  if FAxIdXpgList <> Nil then 
    FreeAndNil(FAxIdXpgList);
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
end;

procedure TCT_RadarChart.Assign(AItem: TCT_RadarChart);
begin
end;

procedure TCT_RadarChart.CopyTo(AItem: TCT_RadarChart);
begin
end;

function  TCT_RadarChart.Create_RadarStyle: TCT_RadarStyle;
begin
  Result := TCT_RadarStyle.Create(FOwner);
  FRadarStyle := Result;
end;

function  TCT_RadarChart.Create_VaryColors: TCT_Boolean;
begin
  Result := TCT_Boolean.Create(FOwner);
  FVaryColors := Result;
end;

function  TCT_RadarChart.Create_SerXpgList: TCT_RadarSerXpgList;
begin
  Result := TCT_RadarSerXpgList.Create(FOwner);
  FSerXpgList := Result;
end;

function  TCT_RadarChart.Create_DLbls: TCT_DLbls;
begin
  Result := TCT_DLbls.Create(FOwner);
  FDLbls := Result;
end;

function  TCT_RadarChart.Create_AxIdXpgList: TCT_UnsignedIntXpgList;
begin
  Result := TCT_UnsignedIntXpgList.Create(FOwner);
  FAxIdXpgList := Result;
end;

function  TCT_RadarChart.Create_ExtLst: TCT_ExtensionList;
begin
  Result := TCT_ExtensionList.Create(FOwner);
  FExtLst := Result;
end;

{ TCT_RadarChartXpgList }

function  TCT_RadarChartXpgList.GetItems(Index: integer): TCT_RadarChart;
begin
  Result := TCT_RadarChart(inherited Items[Index]);
end;

function  TCT_RadarChartXpgList.Add: TCT_RadarChart;
begin
  Result := TCT_RadarChart.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_RadarChartXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_RadarChartXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
    if xaElements in Items[i].Assigneds then
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
end;

procedure TCT_RadarChartXpgList.Assign(AItem: TCT_RadarChartXpgList);
begin
end;

procedure TCT_RadarChartXpgList.CopyTo(AItem: TCT_RadarChartXpgList);
begin
end;

{ TCT_ScatterChart }

function  TCT_ScatterChart.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FScatterStyle <> Nil then 
    Inc(ElemsAssigned,FScatterStyle.CheckAssigned);
  if FVaryColors <> Nil then 
    Inc(ElemsAssigned,FVaryColors.CheckAssigned);
  if FSerXpgList <> Nil then 
    Inc(ElemsAssigned,FSerXpgList.CheckAssigned);
  if FDLbls <> Nil then 
    Inc(ElemsAssigned,FDLbls.CheckAssigned);
  if FAxIdXpgList <> Nil then 
    Inc(ElemsAssigned,FAxIdXpgList.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_ScatterChart.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $000005A4: begin
      if FScatterStyle = Nil then 
        FScatterStyle := TCT_ScatterStyle.Create(FOwner);
      Result := FScatterStyle;
    end;
    $000004D1: begin
      if FVaryColors = Nil then 
        FVaryColors := TCT_Boolean.Create(FOwner);
      Result := FVaryColors;
    end;
    $000001E7: begin
      if FSerXpgList = Nil then 
        FSerXpgList := TCT_ScatterSerXpgList.Create(FOwner);
      Result := FSerXpgList.Add;
    end;
    $0000028E: begin
      if FDLbls = Nil then 
        FDLbls := TCT_DLbls.Create(FOwner);
      Result := FDLbls;
    end;
    $00000223: begin
      if FAxIdXpgList = Nil then 
        FAxIdXpgList := TCT_UnsignedIntXpgList.Create(FOwner);
      Result := FAxIdXpgList.Add;
    end;
    $00000321: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_ScatterChart.Write(AWriter: TXpgWriteXML);
begin
  if (FScatterStyle <> Nil) and FScatterStyle.Assigned then 
  begin
    FScatterStyle.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:scatterStyle');
  end;
  if (FVaryColors <> Nil) and FVaryColors.Assigned then 
  begin
    FVaryColors.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:varyColors');
  end;
  if FSerXpgList <> Nil then 
    FSerXpgList.Write(AWriter,'c:ser');
  if (FDLbls <> Nil) and FDLbls.Assigned then 
    if xaElements in FDLbls.Assigneds then
    begin
      AWriter.BeginTag('c:dLbls');
      FDLbls.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:dLbls');
  if FAxIdXpgList <> Nil then 
    FAxIdXpgList.Write(AWriter,'c:axId');
{$ifdef XLS_WRITE_EXTLST}
  if (FExtLst <> Nil) and FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then
    begin
      AWriter.BeginTag('c:extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:extLst');
{$endif}
end;

constructor TCT_ScatterChart.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 6;
  FAttributeCount := 0;
end;

destructor TCT_ScatterChart.Destroy;
begin
  if FScatterStyle <> Nil then 
    FScatterStyle.Free;
  if FVaryColors <> Nil then 
    FVaryColors.Free;
  if FSerXpgList <> Nil then 
    FSerXpgList.Free;
  if FDLbls <> Nil then 
    FDLbls.Free;
  if FAxIdXpgList <> Nil then 
    FAxIdXpgList.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_ScatterChart.Clear;
begin
  FAssigneds := [];
  if FScatterStyle <> Nil then 
    FreeAndNil(FScatterStyle);
  if FVaryColors <> Nil then 
    FreeAndNil(FVaryColors);
  if FSerXpgList <> Nil then 
    FreeAndNil(FSerXpgList);
  if FDLbls <> Nil then 
    FreeAndNil(FDLbls);
  if FAxIdXpgList <> Nil then 
    FreeAndNil(FAxIdXpgList);
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
end;

procedure TCT_ScatterChart.Assign(AItem: TCT_ScatterChart);
begin
end;

procedure TCT_ScatterChart.CopyTo(AItem: TCT_ScatterChart);
begin
end;

function  TCT_ScatterChart.Create_ScatterStyle: TCT_ScatterStyle;
begin
  Result := TCT_ScatterStyle.Create(FOwner);
  FScatterStyle := Result;
end;

function  TCT_ScatterChart.Create_VaryColors: TCT_Boolean;
begin
  Result := TCT_Boolean.Create(FOwner);
  FVaryColors := Result;
end;

function  TCT_ScatterChart.Create_SerXpgList: TCT_ScatterSerXpgList;
begin
  Result := TCT_ScatterSerXpgList.Create(FOwner);
  FSerXpgList := Result;
end;

function  TCT_ScatterChart.Create_DLbls: TCT_DLbls;
begin
  Result := TCT_DLbls.Create(FOwner);
  FDLbls := Result;
end;

function  TCT_ScatterChart.Create_AxIdXpgList: TCT_UnsignedIntXpgList;
begin
  Result := TCT_UnsignedIntXpgList.Create(FOwner);
  FAxIdXpgList := Result;
end;

function  TCT_ScatterChart.Create_ExtLst: TCT_ExtensionList;
begin
  Result := TCT_ExtensionList.Create(FOwner);
  FExtLst := Result;
end;

{ TCT_ScatterChartXpgList }

function  TCT_ScatterChartXpgList.GetItems(Index: integer): TCT_ScatterChart;
begin
  Result := TCT_ScatterChart(inherited Items[Index]);
end;

function  TCT_ScatterChartXpgList.Add: TCT_ScatterChart;
begin
  Result := TCT_ScatterChart.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_ScatterChartXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_ScatterChartXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
    if xaElements in Items[i].Assigneds then
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
end;

procedure TCT_ScatterChartXpgList.Assign(AItem: TCT_ScatterChartXpgList);
begin
end;

procedure TCT_ScatterChartXpgList.CopyTo(AItem: TCT_ScatterChartXpgList);
begin
end;

{ TCT_PieChart }

function  TCT_PieChart.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FEG_PieChartShared.CheckAssigned);
  if FFirstSliceAng <> Nil then 
    Inc(ElemsAssigned,FFirstSliceAng.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_PieChart.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $000005CB: begin
      if FFirstSliceAng = Nil then 
        FFirstSliceAng := TCT_FirstSliceAng.Create(FOwner);
      Result := FFirstSliceAng;
    end;
    $00000321: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
    begin
      Result := FEG_PieChartShared.HandleElement(AReader);
      if Result = Nil then 
        FOwner.Errors.Error(xemUnknownElement,AReader.QName);
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_PieChart.Write(AWriter: TXpgWriteXML);
begin
  FEG_PieChartShared.Write(AWriter);
  if (FFirstSliceAng <> Nil) and FFirstSliceAng.Assigned then 
  begin
    FFirstSliceAng.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:firstSliceAng');
  end;
{$ifdef XLS_WRITE_EXTLST}
  if (FExtLst <> Nil) and FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then
    begin
      AWriter.BeginTag('c:extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:extLst');
{$endif}
end;

constructor TCT_PieChart.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 3;
  FAttributeCount := 0;
  FEG_PieChartShared := TEG_PieChartShared.Create(FOwner);
end;

destructor TCT_PieChart.Destroy;
begin
  FEG_PieChartShared.Free;
  if FFirstSliceAng <> Nil then 
    FFirstSliceAng.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_PieChart.Clear;
begin
  FAssigneds := [];
  FEG_PieChartShared.Clear;
  if FFirstSliceAng <> Nil then 
    FreeAndNil(FFirstSliceAng);
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
end;

procedure TCT_PieChart.Assign(AItem: TCT_PieChart);
begin
end;

procedure TCT_PieChart.CopyTo(AItem: TCT_PieChart);
begin
end;

function  TCT_PieChart.Create_FirstSliceAng: TCT_FirstSliceAng;
begin
  Result := TCT_FirstSliceAng.Create(FOwner);
  FFirstSliceAng := Result;
end;

function  TCT_PieChart.Create_ExtLst: TCT_ExtensionList;
begin
  Result := TCT_ExtensionList.Create(FOwner);
  FExtLst := Result;
end;

{ TCT_PieChartXpgList }

function  TCT_PieChartXpgList.GetItems(Index: integer): TCT_PieChart;
begin
  Result := TCT_PieChart(inherited Items[Index]);
end;

function  TCT_PieChartXpgList.Add: TCT_PieChart;
begin
  Result := TCT_PieChart.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_PieChartXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_PieChartXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
    if xaElements in Items[i].Assigneds then
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
end;

procedure TCT_PieChartXpgList.Assign(AItem: TCT_PieChartXpgList);
begin
end;

procedure TCT_PieChartXpgList.CopyTo(AItem: TCT_PieChartXpgList);
begin
end;

{ TCT_Pie3DChart }

function  TCT_Pie3DChart.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FEG_PieChartShared.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_Pie3DChart.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $00000321: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
    begin
      Result := FEG_PieChartShared.HandleElement(AReader);
      if Result = Nil then 
        FOwner.Errors.Error(xemUnknownElement,AReader.QName);
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_Pie3DChart.Write(AWriter: TXpgWriteXML);
begin
  FEG_PieChartShared.Write(AWriter);
{$ifdef XLS_WRITE_EXTLST}
  if (FExtLst <> Nil) and FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then
    begin
      AWriter.BeginTag('c:extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:extLst');
{$endif}
end;

constructor TCT_Pie3DChart.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 0;
  FEG_PieChartShared := TEG_PieChartShared.Create(FOwner);
end;

destructor TCT_Pie3DChart.Destroy;
begin
  FEG_PieChartShared.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_Pie3DChart.Clear;
begin
  FAssigneds := [];
  FEG_PieChartShared.Clear;
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
end;

procedure TCT_Pie3DChart.Assign(AItem: TCT_Pie3DChart);
begin
end;

procedure TCT_Pie3DChart.CopyTo(AItem: TCT_Pie3DChart);
begin
end;

function  TCT_Pie3DChart.Create_ExtLst: TCT_ExtensionList;
begin
  Result := TCT_ExtensionList.Create(FOwner);
  FExtLst := Result;
end;

{ TCT_Pie3DChartXpgList }

function  TCT_Pie3DChartXpgList.GetItems(Index: integer): TCT_Pie3DChart;
begin
  Result := TCT_Pie3DChart(inherited Items[Index]);
end;

function  TCT_Pie3DChartXpgList.Add: TCT_Pie3DChart;
begin
  Result := TCT_Pie3DChart.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_Pie3DChartXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_Pie3DChartXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
    if xaElements in Items[i].Assigneds then
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
end;

procedure TCT_Pie3DChartXpgList.Assign(AItem: TCT_Pie3DChartXpgList);
begin
end;

procedure TCT_Pie3DChartXpgList.CopyTo(AItem: TCT_Pie3DChartXpgList);
begin
end;

{ TCT_DoughnutChart }

function  TCT_DoughnutChart.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FEG_PieChartShared.CheckAssigned);
  if FFirstSliceAng <> Nil then 
    Inc(ElemsAssigned,FFirstSliceAng.CheckAssigned);
  if FHoleSize <> Nil then 
    Inc(ElemsAssigned,FHoleSize.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_DoughnutChart.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $000005CB: begin
      if FFirstSliceAng = Nil then 
        FFirstSliceAng := TCT_FirstSliceAng.Create(FOwner);
      Result := FFirstSliceAng;
    end;
    $000003E0: begin
      if FHoleSize = Nil then 
        FHoleSize := TCT_HoleSize.Create(FOwner);
      Result := FHoleSize;
    end;
    $00000321: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
    begin
      Result := FEG_PieChartShared.HandleElement(AReader);
      if Result = Nil then 
        FOwner.Errors.Error(xemUnknownElement,AReader.QName);
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_DoughnutChart.Write(AWriter: TXpgWriteXML);
begin
  FEG_PieChartShared.Write(AWriter);
  if (FFirstSliceAng <> Nil) and FFirstSliceAng.Assigned then 
  begin
    FFirstSliceAng.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:firstSliceAng');
  end;
  if (FHoleSize <> Nil) and FHoleSize.Assigned then 
  begin
    FHoleSize.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:holeSize');
  end;
{$ifdef XLS_WRITE_EXTLST}
  if (FExtLst <> Nil) and FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then
    begin
      AWriter.BeginTag('c:extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:extLst');
{$endif}
end;

constructor TCT_DoughnutChart.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 4;
  FAttributeCount := 0;
  FEG_PieChartShared := TEG_PieChartShared.Create(FOwner);
end;

destructor TCT_DoughnutChart.Destroy;
begin
  FEG_PieChartShared.Free;
  if FFirstSliceAng <> Nil then 
    FFirstSliceAng.Free;
  if FHoleSize <> Nil then 
    FHoleSize.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_DoughnutChart.Clear;
begin
  FAssigneds := [];
  FEG_PieChartShared.Clear;
  if FFirstSliceAng <> Nil then 
    FreeAndNil(FFirstSliceAng);
  if FHoleSize <> Nil then 
    FreeAndNil(FHoleSize);
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
end;

procedure TCT_DoughnutChart.Assign(AItem: TCT_DoughnutChart);
begin
end;

procedure TCT_DoughnutChart.CopyTo(AItem: TCT_DoughnutChart);
begin
end;

function  TCT_DoughnutChart.Create_FirstSliceAng: TCT_FirstSliceAng;
begin
  Result := TCT_FirstSliceAng.Create(FOwner);
  FFirstSliceAng := Result;
end;

function  TCT_DoughnutChart.Create_HoleSize: TCT_HoleSize;
begin
  Result := TCT_HoleSize.Create(FOwner);
  FHoleSize := Result;
end;

function  TCT_DoughnutChart.Create_ExtLst: TCT_ExtensionList;
begin
  Result := TCT_ExtensionList.Create(FOwner);
  FExtLst := Result;
end;

{ TCT_DoughnutChartXpgList }

function  TCT_DoughnutChartXpgList.GetItems(Index: integer): TCT_DoughnutChart;
begin
  Result := TCT_DoughnutChart(inherited Items[Index]);
end;

function  TCT_DoughnutChartXpgList.Add: TCT_DoughnutChart;
begin
  Result := TCT_DoughnutChart.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_DoughnutChartXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_DoughnutChartXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
    if xaElements in Items[i].Assigneds then
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
end;

procedure TCT_DoughnutChartXpgList.Assign(AItem: TCT_DoughnutChartXpgList);
begin
end;

procedure TCT_DoughnutChartXpgList.CopyTo(AItem: TCT_DoughnutChartXpgList);
begin
end;

{ TCT_BarChart }

function  TCT_BarChart.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FEG_BarChartShared.CheckAssigned);
  if FGapWidth <> Nil then 
    Inc(ElemsAssigned,FGapWidth.CheckAssigned);
  if FOverlap <> Nil then 
    Inc(ElemsAssigned,FOverlap.CheckAssigned);
  if FSerLinesXpgList <> Nil then 
    Inc(ElemsAssigned,FSerLinesXpgList.CheckAssigned);
  if FAxIdXpgList <> Nil then 
    Inc(ElemsAssigned,FAxIdXpgList.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_BarChart.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $000003D5: begin
      if FGapWidth = Nil then 
        FGapWidth := TCT_GapAmount.Create(FOwner);
      Result := FGapWidth;
    end;
    $00000396: begin
      if FOverlap = Nil then 
        FOverlap := TCT_Overlap.Create(FOwner);
      Result := FOverlap;
    end;
    $000003E2: begin
      if FSerLinesXpgList = Nil then 
        FSerLinesXpgList := TCT_ChartLinesXpgList.Create(FOwner);
      Result := FSerLinesXpgList.Add;
    end;
    $00000223: begin
      if FAxIdXpgList = Nil then 
        FAxIdXpgList := TCT_UnsignedIntXpgList.Create(FOwner);
      Result := FAxIdXpgList.Add;
    end;
    $00000321: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
    begin
      Result := FEG_BarChartShared.HandleElement(AReader);
      if Result = Nil then 
        FOwner.Errors.Error(xemUnknownElement,AReader.QName);
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_BarChart.Write(AWriter: TXpgWriteXML);
begin
  FEG_BarChartShared.Write(AWriter);
  if (FGapWidth <> Nil) and FGapWidth.Assigned then 
  begin
    FGapWidth.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:gapWidth');
  end;
  if (FOverlap <> Nil) and FOverlap.Assigned then 
  begin
    FOverlap.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:overlap');
  end;
  if FSerLinesXpgList <> Nil then 
    FSerLinesXpgList.Write(AWriter,'c:serLines');
  if FAxIdXpgList <> Nil then 
    FAxIdXpgList.Write(AWriter,'c:axId');
{$ifdef XLS_WRITE_EXTLST}
  if (FExtLst <> Nil) and FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then
    begin
      AWriter.BeginTag('c:extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:extLst');
{$endif}
end;

constructor TCT_BarChart.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 6;
  FAttributeCount := 0;
  FEG_BarChartShared := TEG_BarChartShared.Create(FOwner);
end;

destructor TCT_BarChart.Destroy;
begin
  FEG_BarChartShared.Free;
  if FGapWidth <> Nil then 
    FGapWidth.Free;
  if FOverlap <> Nil then 
    FOverlap.Free;
  if FSerLinesXpgList <> Nil then 
    FSerLinesXpgList.Free;
  if FAxIdXpgList <> Nil then 
    FAxIdXpgList.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_BarChart.Clear;
begin
  FAssigneds := [];
  FEG_BarChartShared.Clear;
  if FGapWidth <> Nil then 
    FreeAndNil(FGapWidth);
  if FOverlap <> Nil then 
    FreeAndNil(FOverlap);
  if FSerLinesXpgList <> Nil then 
    FreeAndNil(FSerLinesXpgList);
  if FAxIdXpgList <> Nil then 
    FreeAndNil(FAxIdXpgList);
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
end;

procedure TCT_BarChart.Assign(AItem: TCT_BarChart);
begin
end;

procedure TCT_BarChart.CopyTo(AItem: TCT_BarChart);
begin
end;

function  TCT_BarChart.Create_GapWidth: TCT_GapAmount;
begin
  Result := TCT_GapAmount.Create(FOwner);
  FGapWidth := Result;
end;

function  TCT_BarChart.Create_Overlap: TCT_Overlap;
begin
  Result := TCT_Overlap.Create(FOwner);
  FOverlap := Result;
end;

function  TCT_BarChart.Create_SerLinesXpgList: TCT_ChartLinesXpgList;
begin
  Result := TCT_ChartLinesXpgList.Create(FOwner);
  FSerLinesXpgList := Result;
end;

function  TCT_BarChart.Create_AxIdXpgList: TCT_UnsignedIntXpgList;
begin
  Result := TCT_UnsignedIntXpgList.Create(FOwner);
  FAxIdXpgList := Result;
end;

function  TCT_BarChart.Create_ExtLst: TCT_ExtensionList;
begin
  Result := TCT_ExtensionList.Create(FOwner);
  FExtLst := Result;
end;

{ TCT_BarChartXpgList }

function  TCT_BarChartXpgList.GetItems(Index: integer): TCT_BarChart;
begin
  Result := TCT_BarChart(inherited Items[Index]);
end;

function  TCT_BarChartXpgList.Add: TCT_BarChart;
begin
  Result := TCT_BarChart.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_BarChartXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_BarChartXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
    if xaElements in Items[i].Assigneds then
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
end;

procedure TCT_BarChartXpgList.Assign(AItem: TCT_BarChartXpgList);
begin
end;

procedure TCT_BarChartXpgList.CopyTo(AItem: TCT_BarChartXpgList);
begin
end;

{ TCT_Bar3DChart }

function  TCT_Bar3DChart.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FEG_BarChartShared.CheckAssigned);
  if FGapWidth <> Nil then 
    Inc(ElemsAssigned,FGapWidth.CheckAssigned);
  if FGapDepth <> Nil then 
    Inc(ElemsAssigned,FGapDepth.CheckAssigned);
  if FShape <> Nil then 
    Inc(ElemsAssigned,FShape.CheckAssigned);
  if FAxIdXpgList <> Nil then 
    Inc(ElemsAssigned,FAxIdXpgList.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_Bar3DChart.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $000003D5: begin
      if FGapWidth = Nil then 
        FGapWidth := TCT_GapAmount.Create(FOwner);
      Result := FGapWidth;
    end;
    $000003CA: begin
      if FGapDepth = Nil then 
        FGapDepth := TCT_GapAmount.Create(FOwner);
      Result := FGapDepth;
    end;
    $000002AE: begin
      if FShape = Nil then 
        FShape := TCT_Shape_Chart.Create(FOwner);
      Result := FShape;
    end;
    $00000223: begin
      if FAxIdXpgList = Nil then 
        FAxIdXpgList := TCT_UnsignedIntXpgList.Create(FOwner);
      Result := FAxIdXpgList.Add;
    end;
    $00000321: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
    begin
      Result := FEG_BarChartShared.HandleElement(AReader);
      if Result = Nil then 
        FOwner.Errors.Error(xemUnknownElement,AReader.QName);
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_Bar3DChart.Write(AWriter: TXpgWriteXML);
begin
  FEG_BarChartShared.Write(AWriter);
  if (FGapWidth <> Nil) and FGapWidth.Assigned then
  begin
    FGapWidth.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:gapWidth');
  end;
  if (FGapDepth <> Nil) and FGapDepth.Assigned then
  begin
    FGapDepth.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:gapDepth');
  end;
  if (FShape <> Nil) and FShape.Assigned then
  begin
    FShape.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:shape');
  end;
  if FAxIdXpgList <> Nil then
    FAxIdXpgList.Write(AWriter,'c:axId');
{$ifdef XLS_WRITE_EXTLST}
  if (FExtLst <> Nil) and FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then
    begin
      AWriter.BeginTag('c:extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:extLst');
{$endif}
end;

constructor TCT_Bar3DChart.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 6;
  FAttributeCount := 0;
  FEG_BarChartShared := TEG_BarChartShared.Create(FOwner);
end;

destructor TCT_Bar3DChart.Destroy;
begin
  FEG_BarChartShared.Free;
  if FGapWidth <> Nil then 
    FGapWidth.Free;
  if FGapDepth <> Nil then 
    FGapDepth.Free;
  if FShape <> Nil then 
    FShape.Free;
  if FAxIdXpgList <> Nil then 
    FAxIdXpgList.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_Bar3DChart.Clear;
begin
  FAssigneds := [];
  FEG_BarChartShared.Clear;
  if FGapWidth <> Nil then 
    FreeAndNil(FGapWidth);
  if FGapDepth <> Nil then 
    FreeAndNil(FGapDepth);
  if FShape <> Nil then 
    FreeAndNil(FShape);
  if FAxIdXpgList <> Nil then 
    FreeAndNil(FAxIdXpgList);
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
end;

procedure TCT_Bar3DChart.Assign(AItem: TCT_Bar3DChart);
begin
end;

procedure TCT_Bar3DChart.CopyTo(AItem: TCT_Bar3DChart);
begin
end;

function  TCT_Bar3DChart.Create_GapWidth: TCT_GapAmount;
begin
  Result := TCT_GapAmount.Create(FOwner);
  FGapWidth := Result;
end;

function  TCT_Bar3DChart.Create_GapDepth: TCT_GapAmount;
begin
  Result := TCT_GapAmount.Create(FOwner);
  FGapDepth := Result;
end;

function  TCT_Bar3DChart.Create_Shape: TCT_Shape_Chart;
begin
  Result := TCT_Shape_Chart.Create(FOwner);
  FShape := Result;
end;

function  TCT_Bar3DChart.Create_AxIdXpgList: TCT_UnsignedIntXpgList;
begin
  Result := TCT_UnsignedIntXpgList.Create(FOwner);
  FAxIdXpgList := Result;
end;

function  TCT_Bar3DChart.Create_ExtLst: TCT_ExtensionList;
begin
  Result := TCT_ExtensionList.Create(FOwner);
  FExtLst := Result;
end;

{ TCT_Bar3DChartXpgList }

function  TCT_Bar3DChartXpgList.GetItems(Index: integer): TCT_Bar3DChart;
begin
  Result := TCT_Bar3DChart(inherited Items[Index]);
end;

function  TCT_Bar3DChartXpgList.Add: TCT_Bar3DChart;
begin
  Result := TCT_Bar3DChart.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_Bar3DChartXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_Bar3DChartXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
    if xaElements in Items[i].Assigneds then
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
end;

procedure TCT_Bar3DChartXpgList.Assign(AItem: TCT_Bar3DChartXpgList);
begin
end;

procedure TCT_Bar3DChartXpgList.CopyTo(AItem: TCT_Bar3DChartXpgList);
begin
end;

{ TCT_OfPieChart }

function  TCT_OfPieChart.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FOfPieType <> Nil then 
    Inc(ElemsAssigned,FOfPieType.CheckAssigned);
  Inc(ElemsAssigned,FEG_PieChartShared.CheckAssigned);
  if FGapWidth <> Nil then 
    Inc(ElemsAssigned,FGapWidth.CheckAssigned);
  if FSplitType <> Nil then 
    Inc(ElemsAssigned,FSplitType.CheckAssigned);
  if FSplitPos <> Nil then 
    Inc(ElemsAssigned,FSplitPos.CheckAssigned);
  if FCustSplit <> Nil then 
    Inc(ElemsAssigned,FCustSplit.CheckAssigned);
  if FSecondPieSize <> Nil then 
    Inc(ElemsAssigned,FSecondPieSize.CheckAssigned);
  if FSerLinesXpgList <> Nil then 
    Inc(ElemsAssigned,FSerLinesXpgList.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_OfPieChart.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $00000432: begin
      if FOfPieType = Nil then 
        FOfPieType := TCT_OfPieType.Create(FOwner);
      Result := FOfPieType;
    end;
    $000003D5: begin
      if FGapWidth = Nil then 
        FGapWidth := TCT_GapAmount.Create(FOwner);
      Result := FGapWidth;
    end;
    $0000046B: begin
      if FSplitType = Nil then 
        FSplitType := TCT_SplitType.Create(FOwner);
      Result := FSplitType;
    end;
    $000003FB: begin
      if FSplitPos = Nil then 
        FSplitPos := TCT_Double.Create(FOwner);
      Result := FSplitPos;
    end;
    $00000468: begin
      if FCustSplit = Nil then 
        FCustSplit := TCT_CustSplit.Create(FOwner);
      Result := FCustSplit;
    end;
    $000005D2: begin
      if FSecondPieSize = Nil then 
        FSecondPieSize := TCT_SecondPieSize.Create(FOwner);
      Result := FSecondPieSize;
    end;
    $000003E2: begin
      if FSerLinesXpgList = Nil then 
        FSerLinesXpgList := TCT_ChartLinesXpgList.Create(FOwner);
      Result := FSerLinesXpgList.Add;
    end;
    $00000321: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
    begin
      Result := FEG_PieChartShared.HandleElement(AReader);
      if Result = Nil then 
        FOwner.Errors.Error(xemUnknownElement,AReader.QName);
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_OfPieChart.Write(AWriter: TXpgWriteXML);
begin
  if (FOfPieType <> Nil) and FOfPieType.Assigned then 
  begin
    FOfPieType.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:ofPieType');
  end;
  FEG_PieChartShared.Write(AWriter);
  if (FGapWidth <> Nil) and FGapWidth.Assigned then 
  begin
    FGapWidth.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:gapWidth');
  end;
  if (FSplitType <> Nil) and FSplitType.Assigned then 
  begin
    FSplitType.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:splitType');
  end;
  if (FSplitPos <> Nil) and FSplitPos.Assigned then 
  begin
    FSplitPos.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:splitPos');
  end;
  if (FCustSplit <> Nil) and FCustSplit.Assigned then 
    if xaElements in FCustSplit.Assigneds then
    begin
      AWriter.BeginTag('c:custSplit');
      FCustSplit.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:custSplit');
  if (FSecondPieSize <> Nil) and FSecondPieSize.Assigned then 
  begin
    FSecondPieSize.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:secondPieSize');
  end;
  if FSerLinesXpgList <> Nil then 
    FSerLinesXpgList.Write(AWriter,'c:serLines');
{$ifdef XLS_WRITE_EXTLST}
  if (FExtLst <> Nil) and FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then
    begin
      AWriter.BeginTag('c:extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:extLst');
{$endif}
end;

constructor TCT_OfPieChart.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 9;
  FAttributeCount := 0;
  FEG_PieChartShared := TEG_PieChartShared.Create(FOwner);
end;

destructor TCT_OfPieChart.Destroy;
begin
  if FOfPieType <> Nil then 
    FOfPieType.Free;
  FEG_PieChartShared.Free;
  if FGapWidth <> Nil then 
    FGapWidth.Free;
  if FSplitType <> Nil then 
    FSplitType.Free;
  if FSplitPos <> Nil then 
    FSplitPos.Free;
  if FCustSplit <> Nil then 
    FCustSplit.Free;
  if FSecondPieSize <> Nil then 
    FSecondPieSize.Free;
  if FSerLinesXpgList <> Nil then 
    FSerLinesXpgList.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_OfPieChart.Clear;
begin
  FAssigneds := [];
  if FOfPieType <> Nil then 
    FreeAndNil(FOfPieType);
  FEG_PieChartShared.Clear;
  if FGapWidth <> Nil then 
    FreeAndNil(FGapWidth);
  if FSplitType <> Nil then 
    FreeAndNil(FSplitType);
  if FSplitPos <> Nil then 
    FreeAndNil(FSplitPos);
  if FCustSplit <> Nil then 
    FreeAndNil(FCustSplit);
  if FSecondPieSize <> Nil then 
    FreeAndNil(FSecondPieSize);
  if FSerLinesXpgList <> Nil then 
    FreeAndNil(FSerLinesXpgList);
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
end;

procedure TCT_OfPieChart.Assign(AItem: TCT_OfPieChart);
begin
end;

procedure TCT_OfPieChart.CopyTo(AItem: TCT_OfPieChart);
begin
end;

function  TCT_OfPieChart.Create_OfPieType: TCT_OfPieType;
begin
  Result := TCT_OfPieType.Create(FOwner);
  FOfPieType := Result;
end;

function  TCT_OfPieChart.Create_GapWidth: TCT_GapAmount;
begin
  Result := TCT_GapAmount.Create(FOwner);
  FGapWidth := Result;
end;

function  TCT_OfPieChart.Create_SplitType: TCT_SplitType;
begin
  Result := TCT_SplitType.Create(FOwner);
  FSplitType := Result;
end;

function  TCT_OfPieChart.Create_SplitPos: TCT_Double;
begin
  Result := TCT_Double.Create(FOwner);
  FSplitPos := Result;
end;

function  TCT_OfPieChart.Create_CustSplit: TCT_CustSplit;
begin
  Result := TCT_CustSplit.Create(FOwner);
  FCustSplit := Result;
end;

function  TCT_OfPieChart.Create_SecondPieSize: TCT_SecondPieSize;
begin
  Result := TCT_SecondPieSize.Create(FOwner);
  FSecondPieSize := Result;
end;

function  TCT_OfPieChart.Create_SerLinesXpgList: TCT_ChartLinesXpgList;
begin
  Result := TCT_ChartLinesXpgList.Create(FOwner);
  FSerLinesXpgList := Result;
end;

function  TCT_OfPieChart.Create_ExtLst: TCT_ExtensionList;
begin
  Result := TCT_ExtensionList.Create(FOwner);
  FExtLst := Result;
end;

{ TCT_OfPieChartXpgList }

function  TCT_OfPieChartXpgList.GetItems(Index: integer): TCT_OfPieChart;
begin
  Result := TCT_OfPieChart(inherited Items[Index]);
end;

function  TCT_OfPieChartXpgList.Add: TCT_OfPieChart;
begin
  Result := TCT_OfPieChart.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_OfPieChartXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_OfPieChartXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
    if xaElements in Items[i].Assigneds then
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
end;

procedure TCT_OfPieChartXpgList.Assign(AItem: TCT_OfPieChartXpgList);
begin
end;

procedure TCT_OfPieChartXpgList.CopyTo(AItem: TCT_OfPieChartXpgList);
begin
end;

{ TCT_SurfaceChart }

function  TCT_SurfaceChart.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FEG_SurfaceChartShared.CheckAssigned);
  if FAxIdXpgList <> Nil then 
    Inc(ElemsAssigned,FAxIdXpgList.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_SurfaceChart.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $00000223: begin
      if FAxIdXpgList = Nil then 
        FAxIdXpgList := TCT_UnsignedIntXpgList.Create(FOwner);
      Result := FAxIdXpgList.Add;
    end;
    $00000321: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
    begin
      Result := FEG_SurfaceChartShared.HandleElement(AReader);
      if Result = Nil then 
        FOwner.Errors.Error(xemUnknownElement,AReader.QName);
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_SurfaceChart.Write(AWriter: TXpgWriteXML);
begin
  FEG_SurfaceChartShared.Write(AWriter);
  if FAxIdXpgList <> Nil then 
    FAxIdXpgList.Write(AWriter,'c:axId');
{$ifdef XLS_WRITE_EXTLST}
  if (FExtLst <> Nil) and FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then
    begin
      AWriter.BeginTag('c:extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:extLst');
{$endif}
end;

constructor TCT_SurfaceChart.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 3;
  FAttributeCount := 0;
  FEG_SurfaceChartShared := TEG_SurfaceChartShared.Create(FOwner);
end;

destructor TCT_SurfaceChart.Destroy;
begin
  FEG_SurfaceChartShared.Free;
  if FAxIdXpgList <> Nil then 
    FAxIdXpgList.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_SurfaceChart.Clear;
begin
  FAssigneds := [];
  FEG_SurfaceChartShared.Clear;
  if FAxIdXpgList <> Nil then 
    FreeAndNil(FAxIdXpgList);
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
end;

procedure TCT_SurfaceChart.Assign(AItem: TCT_SurfaceChart);
begin
end;

procedure TCT_SurfaceChart.CopyTo(AItem: TCT_SurfaceChart);
begin
end;

function  TCT_SurfaceChart.Create_AxIdXpgList: TCT_UnsignedIntXpgList;
begin
  Result := TCT_UnsignedIntXpgList.Create(FOwner);
  FAxIdXpgList := Result;
end;

function  TCT_SurfaceChart.Create_ExtLst: TCT_ExtensionList;
begin
  Result := TCT_ExtensionList.Create(FOwner);
  FExtLst := Result;
end;

{ TCT_SurfaceChartXpgList }

function  TCT_SurfaceChartXpgList.GetItems(Index: integer): TCT_SurfaceChart;
begin
  Result := TCT_SurfaceChart(inherited Items[Index]);
end;

function  TCT_SurfaceChartXpgList.Add: TCT_SurfaceChart;
begin
  Result := TCT_SurfaceChart.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_SurfaceChartXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_SurfaceChartXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
    if xaElements in Items[i].Assigneds then
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
end;

procedure TCT_SurfaceChartXpgList.Assign(AItem: TCT_SurfaceChartXpgList);
begin
end;

procedure TCT_SurfaceChartXpgList.CopyTo(AItem: TCT_SurfaceChartXpgList);
begin
end;

{ TCT_Surface3DChart }

function  TCT_Surface3DChart.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FEG_SurfaceChartShared.CheckAssigned);
  if FAxIdXpgList <> Nil then 
    Inc(ElemsAssigned,FAxIdXpgList.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_Surface3DChart.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $00000223: begin
      if FAxIdXpgList = Nil then 
        FAxIdXpgList := TCT_UnsignedIntXpgList.Create(FOwner);
      Result := FAxIdXpgList.Add;
    end;
    $00000321: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
    begin
      Result := FEG_SurfaceChartShared.HandleElement(AReader);
      if Result = Nil then 
        FOwner.Errors.Error(xemUnknownElement,AReader.QName);
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_Surface3DChart.Write(AWriter: TXpgWriteXML);
begin
  FEG_SurfaceChartShared.Write(AWriter);
  if FAxIdXpgList <> Nil then 
    FAxIdXpgList.Write(AWriter,'c:axId');
{$ifdef XLS_WRITE_EXTLST}
  if (FExtLst <> Nil) and FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then
    begin
      AWriter.BeginTag('c:extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:extLst');
{$endif}
end;

constructor TCT_Surface3DChart.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 3;
  FAttributeCount := 0;
  FEG_SurfaceChartShared := TEG_SurfaceChartShared.Create(FOwner);
end;

destructor TCT_Surface3DChart.Destroy;
begin
  FEG_SurfaceChartShared.Free;
  if FAxIdXpgList <> Nil then 
    FAxIdXpgList.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_Surface3DChart.Clear;
begin
  FAssigneds := [];
  FEG_SurfaceChartShared.Clear;
  if FAxIdXpgList <> Nil then 
    FreeAndNil(FAxIdXpgList);
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
end;

procedure TCT_Surface3DChart.Assign(AItem: TCT_Surface3DChart);
begin
end;

procedure TCT_Surface3DChart.CopyTo(AItem: TCT_Surface3DChart);
begin
end;

function  TCT_Surface3DChart.Create_AxIdXpgList: TCT_UnsignedIntXpgList;
begin
  Result := TCT_UnsignedIntXpgList.Create(FOwner);
  FAxIdXpgList := Result;
end;

function  TCT_Surface3DChart.Create_ExtLst: TCT_ExtensionList;
begin
  Result := TCT_ExtensionList.Create(FOwner);
  FExtLst := Result;
end;

{ TCT_Surface3DChartXpgList }

function  TCT_Surface3DChartXpgList.GetItems(Index: integer): TCT_Surface3DChart;
begin
  Result := TCT_Surface3DChart(inherited Items[Index]);
end;

function  TCT_Surface3DChartXpgList.Add: TCT_Surface3DChart;
begin
  Result := TCT_Surface3DChart.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_Surface3DChartXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_Surface3DChartXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
    if xaElements in Items[i].Assigneds then
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
end;

procedure TCT_Surface3DChartXpgList.Assign(AItem: TCT_Surface3DChartXpgList);
begin
end;

procedure TCT_Surface3DChartXpgList.CopyTo(AItem: TCT_Surface3DChartXpgList);
begin
end;

{ TCT_BubbleChart }

function  TCT_BubbleChart.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FVaryColors <> Nil then 
    Inc(ElemsAssigned,FVaryColors.CheckAssigned);
  if FSerXpgList <> Nil then 
    Inc(ElemsAssigned,FSerXpgList.CheckAssigned);
  if FDLbls <> Nil then 
    Inc(ElemsAssigned,FDLbls.CheckAssigned);
  if FBubble3D <> Nil then 
    Inc(ElemsAssigned,FBubble3D.CheckAssigned);
  if FBubbleScale <> Nil then 
    Inc(ElemsAssigned,FBubbleScale.CheckAssigned);
  if FShowNegBubbles <> Nil then 
    Inc(ElemsAssigned,FShowNegBubbles.CheckAssigned);
  if FSizeRepresents <> Nil then 
    Inc(ElemsAssigned,FSizeRepresents.CheckAssigned);
  if FAxIdXpgList <> Nil then 
    Inc(ElemsAssigned,FAxIdXpgList.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_BubbleChart.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $000004D1: begin
      if FVaryColors = Nil then 
        FVaryColors := TCT_Boolean.Create(FOwner);
      Result := FVaryColors;
    end;
    $000001E7: begin
      if FSerXpgList = Nil then 
        FSerXpgList := TCT_BubbleSerXpgList.Create(FOwner);
      Result := FSerXpgList.Add;
    end;
    $0000028E: begin
      if FDLbls = Nil then 
        FDLbls := TCT_DLbls.Create(FOwner);
      Result := FDLbls;
    end;
    $00000380: begin
      if FBubble3D = Nil then 
        FBubble3D := TCT_Boolean.Create(FOwner);
      Result := FBubble3D;
    end;
    $000004F1: begin
      if FBubbleScale = Nil then 
        FBubbleScale := TCT_BubbleScale.Create(FOwner);
      Result := FBubbleScale;
    end;
    $00000637: begin
      if FShowNegBubbles = Nil then 
        FShowNegBubbles := TCT_Boolean.Create(FOwner);
      Result := FShowNegBubbles;
    end;
    $00000683: begin
      if FSizeRepresents = Nil then 
        FSizeRepresents := TCT_SizeRepresents.Create(FOwner);
      Result := FSizeRepresents;
    end;
    $00000223: begin
      if FAxIdXpgList = Nil then 
        FAxIdXpgList := TCT_UnsignedIntXpgList.Create(FOwner);
      Result := FAxIdXpgList.Add;
    end;
    $00000321: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_BubbleChart.Write(AWriter: TXpgWriteXML);
begin
  if (FVaryColors <> Nil) and FVaryColors.Assigned then 
  begin
    FVaryColors.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:varyColors');
  end;
  if FSerXpgList <> Nil then 
    FSerXpgList.Write(AWriter,'c:ser');
  if (FDLbls <> Nil) and FDLbls.Assigned then 
    if xaElements in FDLbls.Assigneds then
    begin
      AWriter.BeginTag('c:dLbls');
      FDLbls.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:dLbls');
  if (FBubble3D <> Nil) and FBubble3D.Assigned then 
  begin
    FBubble3D.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:bubble3D');
  end;
  if (FBubbleScale <> Nil) and FBubbleScale.Assigned then 
  begin
    FBubbleScale.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:bubbleScale');
  end;
  if (FShowNegBubbles <> Nil) and FShowNegBubbles.Assigned then 
  begin
    FShowNegBubbles.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:showNegBubbles');
  end;
  if (FSizeRepresents <> Nil) and FSizeRepresents.Assigned then 
  begin
    FSizeRepresents.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:sizeRepresents');
  end;
  if FAxIdXpgList <> Nil then 
    FAxIdXpgList.Write(AWriter,'c:axId');
{$ifdef XLS_WRITE_EXTLST}
  if (FExtLst <> Nil) and FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then
    begin
      AWriter.BeginTag('c:extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:extLst');
{$endif}
end;

constructor TCT_BubbleChart.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 9;
  FAttributeCount := 0;
end;

destructor TCT_BubbleChart.Destroy;
begin
  if FVaryColors <> Nil then 
    FVaryColors.Free;
  if FSerXpgList <> Nil then 
    FSerXpgList.Free;
  if FDLbls <> Nil then 
    FDLbls.Free;
  if FBubble3D <> Nil then 
    FBubble3D.Free;
  if FBubbleScale <> Nil then 
    FBubbleScale.Free;
  if FShowNegBubbles <> Nil then 
    FShowNegBubbles.Free;
  if FSizeRepresents <> Nil then 
    FSizeRepresents.Free;
  if FAxIdXpgList <> Nil then 
    FAxIdXpgList.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_BubbleChart.Clear;
begin
  FAssigneds := [];
  if FVaryColors <> Nil then 
    FreeAndNil(FVaryColors);
  if FSerXpgList <> Nil then 
    FreeAndNil(FSerXpgList);
  if FDLbls <> Nil then 
    FreeAndNil(FDLbls);
  if FBubble3D <> Nil then 
    FreeAndNil(FBubble3D);
  if FBubbleScale <> Nil then 
    FreeAndNil(FBubbleScale);
  if FShowNegBubbles <> Nil then 
    FreeAndNil(FShowNegBubbles);
  if FSizeRepresents <> Nil then 
    FreeAndNil(FSizeRepresents);
  if FAxIdXpgList <> Nil then 
    FreeAndNil(FAxIdXpgList);
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
end;

procedure TCT_BubbleChart.Assign(AItem: TCT_BubbleChart);
begin
end;

procedure TCT_BubbleChart.CopyTo(AItem: TCT_BubbleChart);
begin
end;

function  TCT_BubbleChart.Create_VaryColors: TCT_Boolean;
begin
  Result := TCT_Boolean.Create(FOwner);
  FVaryColors := Result;
end;

function  TCT_BubbleChart.Create_SerXpgList: TCT_BubbleSerXpgList;
begin
  Result := TCT_BubbleSerXpgList.Create(FOwner);
  FSerXpgList := Result;
end;

function  TCT_BubbleChart.Create_DLbls: TCT_DLbls;
begin
  Result := TCT_DLbls.Create(FOwner);
  FDLbls := Result;
end;

function  TCT_BubbleChart.Create_Bubble3D: TCT_Boolean;
begin
  Result := TCT_Boolean.Create(FOwner);
  FBubble3D := Result;
end;

function  TCT_BubbleChart.Create_BubbleScale: TCT_BubbleScale;
begin
  Result := TCT_BubbleScale.Create(FOwner);
  FBubbleScale := Result;
end;

function  TCT_BubbleChart.Create_ShowNegBubbles: TCT_Boolean;
begin
  Result := TCT_Boolean.Create(FOwner);
  FShowNegBubbles := Result;
end;

function  TCT_BubbleChart.Create_SizeRepresents: TCT_SizeRepresents;
begin
  Result := TCT_SizeRepresents.Create(FOwner);
  FSizeRepresents := Result;
end;

function  TCT_BubbleChart.Create_AxIdXpgList: TCT_UnsignedIntXpgList;
begin
  Result := TCT_UnsignedIntXpgList.Create(FOwner);
  FAxIdXpgList := Result;
end;

function  TCT_BubbleChart.Create_ExtLst: TCT_ExtensionList;
begin
  Result := TCT_ExtensionList.Create(FOwner);
  FExtLst := Result;
end;

{ TCT_BubbleChartXpgList }

function  TCT_BubbleChartXpgList.GetItems(Index: integer): TCT_BubbleChart;
begin
  Result := TCT_BubbleChart(inherited Items[Index]);
end;

function  TCT_BubbleChartXpgList.Add: TCT_BubbleChart;
begin
  Result := TCT_BubbleChart.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_BubbleChartXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_BubbleChartXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
    if xaElements in Items[i].Assigneds then
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
end;

procedure TCT_BubbleChartXpgList.Assign(AItem: TCT_BubbleChartXpgList);
begin
end;

procedure TCT_BubbleChartXpgList.CopyTo(AItem: TCT_BubbleChartXpgList);
begin
end;

{ TCT_ValAx }

function  TCT_ValAx.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FEG_AxShared.CheckAssigned);
  if FCrossBetween <> Nil then 
    Inc(ElemsAssigned,FCrossBetween.CheckAssigned);
  if FMajorUnit <> Nil then 
    Inc(ElemsAssigned,FMajorUnit.CheckAssigned);
  if FMinorUnit <> Nil then 
    Inc(ElemsAssigned,FMinorUnit.CheckAssigned);
  if FDispUnits <> Nil then 
    Inc(ElemsAssigned,FDispUnits.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_ValAx.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $00000591: begin
      if FCrossBetween = Nil then 
        FCrossBetween := TCT_CrossBetween.Create(FOwner);
      Result := FCrossBetween;
    end;
    $00000456: begin
      if FMajorUnit = Nil then 
        FMajorUnit := TCT_AxisUnit.Create(FOwner);
      Result := FMajorUnit;
    end;
    $00000462: begin
      if FMinorUnit = Nil then 
        FMinorUnit := TCT_AxisUnit.Create(FOwner);
      Result := FMinorUnit;
    end;
    $00000460: begin
      if FDispUnits = Nil then 
        FDispUnits := TCT_DispUnits.Create(FOwner);
      Result := FDispUnits;
    end;
    $00000321: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
    begin
      Result := FEG_AxShared.HandleElement(AReader);
      if Result = Nil then 
        FOwner.Errors.Error(xemUnknownElement,AReader.QName);
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_ValAx.Write(AWriter: TXpgWriteXML);
begin
  FEG_AxShared.Write(AWriter);
  if (FCrossBetween <> Nil) and FCrossBetween.Assigned then 
  begin
    FCrossBetween.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:crossBetween');
  end;
  if (FMajorUnit <> Nil) and FMajorUnit.Assigned then 
  begin
    FMajorUnit.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:majorUnit');
  end;
  if (FMinorUnit <> Nil) and FMinorUnit.Assigned then 
  begin
    FMinorUnit.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:minorUnit');
  end;
  if (FDispUnits <> Nil) and FDispUnits.Assigned then 
    if xaElements in FDispUnits.Assigneds then
    begin
      AWriter.BeginTag('c:dispUnits');
      FDispUnits.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:dispUnits');
{$ifdef XLS_WRITE_EXTLST}
  if (FExtLst <> Nil) and FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then
    begin
      AWriter.BeginTag('c:extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:extLst');
{$endif}
end;

constructor TCT_ValAx.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 6;
  FAttributeCount := 0;
  FEG_AxShared := TEG_AxShared.Create(FOwner);
end;

destructor TCT_ValAx.Destroy;
begin
  FEG_AxShared.Free;
  if FCrossBetween <> Nil then 
    FCrossBetween.Free;
  if FMajorUnit <> Nil then 
    FMajorUnit.Free;
  if FMinorUnit <> Nil then 
    FMinorUnit.Free;
  if FDispUnits <> Nil then 
    FDispUnits.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_ValAx.Clear;
begin
  FAssigneds := [];
  FEG_AxShared.Clear;
  if FCrossBetween <> Nil then 
    FreeAndNil(FCrossBetween);
  if FMajorUnit <> Nil then 
    FreeAndNil(FMajorUnit);
  if FMinorUnit <> Nil then 
    FreeAndNil(FMinorUnit);
  if FDispUnits <> Nil then 
    FreeAndNil(FDispUnits);
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
end;

procedure TCT_ValAx.Assign(AItem: TCT_ValAx);
begin
end;

procedure TCT_ValAx.CopyTo(AItem: TCT_ValAx);
begin
end;

function  TCT_ValAx.Create_CrossBetween: TCT_CrossBetween;
begin
  Result := TCT_CrossBetween.Create(FOwner);
  FCrossBetween := Result;
end;

function  TCT_ValAx.Create_MajorUnit: TCT_AxisUnit;
begin
  Result := TCT_AxisUnit.Create(FOwner);
  FMajorUnit := Result;
end;

function  TCT_ValAx.Create_MinorUnit: TCT_AxisUnit;
begin
  Result := TCT_AxisUnit.Create(FOwner);
  FMinorUnit := Result;
end;

function  TCT_ValAx.Create_DispUnits: TCT_DispUnits;
begin
  Result := TCT_DispUnits.Create(FOwner);
  FDispUnits := Result;
end;

function  TCT_ValAx.Create_ExtLst: TCT_ExtensionList;
begin
  Result := TCT_ExtensionList.Create(FOwner);
  FExtLst := Result;
end;

{ TCT_ValAxXpgList }

function  TCT_ValAxXpgList.GetItems(Index: integer): TCT_ValAx;
begin
  Result := TCT_ValAx(inherited Items[Index]);
end;

function  TCT_ValAxXpgList.Add: TCT_ValAx;
begin
  Result := TCT_ValAx.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_ValAxXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_ValAxXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
    if xaElements in Items[i].Assigneds then
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
end;

procedure TCT_ValAxXpgList.Assign(AItem: TCT_ValAxXpgList);
begin
end;

procedure TCT_ValAxXpgList.CopyTo(AItem: TCT_ValAxXpgList);
begin
end;

{ TCT_CatAx }

function  TCT_CatAx.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FEG_AxShared.CheckAssigned);
  if FAuto <> Nil then 
    Inc(ElemsAssigned,FAuto.CheckAssigned);
  if FLblAlgn <> Nil then 
    Inc(ElemsAssigned,FLblAlgn.CheckAssigned);
  if FLblOffset <> Nil then 
    Inc(ElemsAssigned,FLblOffset.CheckAssigned);
  if FTickLblSkip <> Nil then 
    Inc(ElemsAssigned,FTickLblSkip.CheckAssigned);
  if FTickMarkSkip <> Nil then 
    Inc(ElemsAssigned,FTickMarkSkip.CheckAssigned);
  if FNoMultiLvlLbl <> Nil then 
    Inc(ElemsAssigned,FNoMultiLvlLbl.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_CatAx.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $00000256: begin
      if FAuto = Nil then 
        FAuto := TCT_Boolean.Create(FOwner);
      Result := FAuto;
    end;
    $00000359: begin
      if FLblAlgn = Nil then 
        FLblAlgn := TCT_LblAlgn.Create(FOwner);
      Result := FLblAlgn;
    end;
    $0000043E: begin
      if FLblOffset = Nil then 
        FLblOffset := TCT_LblOffset.Create(FOwner);
      Result := FLblOffset;
    end;
    $000004F9: begin
      if FTickLblSkip = Nil then 
        FTickLblSkip := TCT_Skip.Create(FOwner);
      Result := FTickLblSkip;
    end;
    $0000056A: begin
      if FTickMarkSkip = Nil then 
        FTickMarkSkip := TCT_Skip.Create(FOwner);
      Result := FTickMarkSkip;
    end;
    $000005CD: begin
      if FNoMultiLvlLbl = Nil then 
        FNoMultiLvlLbl := TCT_Boolean.Create(FOwner);
      Result := FNoMultiLvlLbl;
    end;
    $00000321: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
    begin
      Result := FEG_AxShared.HandleElement(AReader);
      if Result = Nil then 
        FOwner.Errors.Error(xemUnknownElement,AReader.QName);
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_CatAx.Write(AWriter: TXpgWriteXML);
begin
  FEG_AxShared.Write(AWriter);
  if (FAuto <> Nil) and FAuto.Assigned then 
  begin
    FAuto.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:auto');
  end;
  if (FLblAlgn <> Nil) and FLblAlgn.Assigned then 
  begin
    FLblAlgn.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:lblAlgn');
  end;
  if (FLblOffset <> Nil) and FLblOffset.Assigned then 
  begin
    FLblOffset.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:lblOffset');
  end;
  if (FTickLblSkip <> Nil) and FTickLblSkip.Assigned then 
  begin
    FTickLblSkip.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:tickLblSkip');
  end;
  if (FTickMarkSkip <> Nil) and FTickMarkSkip.Assigned then 
  begin
    FTickMarkSkip.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:tickMarkSkip');
  end;
  if (FNoMultiLvlLbl <> Nil) and FNoMultiLvlLbl.Assigned then 
  begin
    FNoMultiLvlLbl.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:noMultiLvlLbl');
  end;
{$ifdef XLS_WRITE_EXTLST}
  if (FExtLst <> Nil) and FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then
    begin
      AWriter.BeginTag('c:extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:extLst');
{$endif}
end;

constructor TCT_CatAx.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 8;
  FAttributeCount := 0;
  FEG_AxShared := TEG_AxShared.Create(FOwner);
end;

destructor TCT_CatAx.Destroy;
begin
  FEG_AxShared.Free;
  if FAuto <> Nil then 
    FAuto.Free;
  if FLblAlgn <> Nil then 
    FLblAlgn.Free;
  if FLblOffset <> Nil then 
    FLblOffset.Free;
  if FTickLblSkip <> Nil then 
    FTickLblSkip.Free;
  if FTickMarkSkip <> Nil then 
    FTickMarkSkip.Free;
  if FNoMultiLvlLbl <> Nil then 
    FNoMultiLvlLbl.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_CatAx.Clear;
begin
  FAssigneds := [];
  FEG_AxShared.Clear;
  if FAuto <> Nil then 
    FreeAndNil(FAuto);
  if FLblAlgn <> Nil then 
    FreeAndNil(FLblAlgn);
  if FLblOffset <> Nil then 
    FreeAndNil(FLblOffset);
  if FTickLblSkip <> Nil then 
    FreeAndNil(FTickLblSkip);
  if FTickMarkSkip <> Nil then 
    FreeAndNil(FTickMarkSkip);
  if FNoMultiLvlLbl <> Nil then 
    FreeAndNil(FNoMultiLvlLbl);
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
end;

procedure TCT_CatAx.Assign(AItem: TCT_CatAx);
begin
end;

procedure TCT_CatAx.CopyTo(AItem: TCT_CatAx);
begin
end;

function  TCT_CatAx.Create_Auto: TCT_Boolean;
begin
  Result := TCT_Boolean.Create(FOwner);
  FAuto := Result;
end;

function  TCT_CatAx.Create_LblAlgn: TCT_LblAlgn;
begin
  Result := TCT_LblAlgn.Create(FOwner);
  FLblAlgn := Result;
end;

function  TCT_CatAx.Create_LblOffset: TCT_LblOffset;
begin
  Result := TCT_LblOffset.Create(FOwner);
  FLblOffset := Result;
end;

function  TCT_CatAx.Create_TickLblSkip: TCT_Skip;
begin
  Result := TCT_Skip.Create(FOwner);
  FTickLblSkip := Result;
end;

function  TCT_CatAx.Create_TickMarkSkip: TCT_Skip;
begin
  Result := TCT_Skip.Create(FOwner);
  FTickMarkSkip := Result;
end;

function  TCT_CatAx.Create_NoMultiLvlLbl: TCT_Boolean;
begin
  Result := TCT_Boolean.Create(FOwner);
  FNoMultiLvlLbl := Result;
end;

function  TCT_CatAx.Create_ExtLst: TCT_ExtensionList;
begin
  Result := TCT_ExtensionList.Create(FOwner);
  FExtLst := Result;
end;

{ TCT_CatAxXpgList }

function  TCT_CatAxXpgList.GetItems(Index: integer): TCT_CatAx;
begin
  Result := TCT_CatAx(inherited Items[Index]);
end;

function  TCT_CatAxXpgList.Add: TCT_CatAx;
begin
  Result := TCT_CatAx.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_CatAxXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_CatAxXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
    if xaElements in Items[i].Assigneds then
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
end;

procedure TCT_CatAxXpgList.Assign(AItem: TCT_CatAxXpgList);
begin
end;

procedure TCT_CatAxXpgList.CopyTo(AItem: TCT_CatAxXpgList);
begin
end;

{ TCT_DateAx }

function  TCT_DateAx.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FEG_AxShared.CheckAssigned);
  if FAuto <> Nil then 
    Inc(ElemsAssigned,FAuto.CheckAssigned);
  if FLblOffset <> Nil then 
    Inc(ElemsAssigned,FLblOffset.CheckAssigned);
  if FBaseTimeUnit <> Nil then 
    Inc(ElemsAssigned,FBaseTimeUnit.CheckAssigned);
  if FMajorUnit <> Nil then 
    Inc(ElemsAssigned,FMajorUnit.CheckAssigned);
  if FMajorTimeUnit <> Nil then 
    Inc(ElemsAssigned,FMajorTimeUnit.CheckAssigned);
  if FMinorUnit <> Nil then 
    Inc(ElemsAssigned,FMinorUnit.CheckAssigned);
  if FMinorTimeUnit <> Nil then 
    Inc(ElemsAssigned,FMinorTimeUnit.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_DateAx.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $00000256: begin
      if FAuto = Nil then 
        FAuto := TCT_Boolean.Create(FOwner);
      Result := FAuto;
    end;
    $0000043E: begin
      if FLblOffset = Nil then 
        FLblOffset := TCT_LblOffset.Create(FOwner);
      Result := FLblOffset;
    end;
    $00000567: begin
      if FBaseTimeUnit = Nil then 
        FBaseTimeUnit := TCT_TimeUnit.Create(FOwner);
      Result := FBaseTimeUnit;
    end;
    $00000456: begin
      if FMajorUnit = Nil then 
        FMajorUnit := TCT_AxisUnit.Create(FOwner);
      Result := FMajorUnit;
    end;
    $000005E5: begin
      if FMajorTimeUnit = Nil then 
        FMajorTimeUnit := TCT_TimeUnit.Create(FOwner);
      Result := FMajorTimeUnit;
    end;
    $00000462: begin
      if FMinorUnit = Nil then 
        FMinorUnit := TCT_AxisUnit.Create(FOwner);
      Result := FMinorUnit;
    end;
    $000005F1: begin
      if FMinorTimeUnit = Nil then 
        FMinorTimeUnit := TCT_TimeUnit.Create(FOwner);
      Result := FMinorTimeUnit;
    end;
    $00000321: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
    begin
      Result := FEG_AxShared.HandleElement(AReader);
      if Result = Nil then 
        FOwner.Errors.Error(xemUnknownElement,AReader.QName);
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_DateAx.Write(AWriter: TXpgWriteXML);
begin
  FEG_AxShared.Write(AWriter);
  if (FAuto <> Nil) and FAuto.Assigned then 
  begin
    FAuto.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:auto');
  end;
//  if (FLblOffset <> Nil) and FLblOffset.Assigned then
//  begin
    FLblOffset.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:lblOffset');
//  end;
  if (FBaseTimeUnit <> Nil) and FBaseTimeUnit.Assigned then 
  begin
    FBaseTimeUnit.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:baseTimeUnit');
  end;
  if (FMajorUnit <> Nil) and FMajorUnit.Assigned then 
  begin
    FMajorUnit.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:majorUnit');
  end;
  if (FMajorTimeUnit <> Nil) and FMajorTimeUnit.Assigned then 
  begin
    FMajorTimeUnit.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:majorTimeUnit');
  end;
  if (FMinorUnit <> Nil) and FMinorUnit.Assigned then 
  begin
    FMinorUnit.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:minorUnit');
  end;
  if (FMinorTimeUnit <> Nil) and FMinorTimeUnit.Assigned then 
  begin
    FMinorTimeUnit.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:minorTimeUnit');
  end;
{$ifdef XLS_WRITE_EXTLST}
  if (FExtLst <> Nil) and FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then
    begin
      AWriter.BeginTag('c:extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:extLst');
{$endif}
end;

constructor TCT_DateAx.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 9;
  FAttributeCount := 0;
  FEG_AxShared := TEG_AxShared.Create(FOwner);
end;

destructor TCT_DateAx.Destroy;
begin
  FEG_AxShared.Free;
  if FAuto <> Nil then 
    FAuto.Free;
  if FLblOffset <> Nil then 
    FLblOffset.Free;
  if FBaseTimeUnit <> Nil then 
    FBaseTimeUnit.Free;
  if FMajorUnit <> Nil then 
    FMajorUnit.Free;
  if FMajorTimeUnit <> Nil then 
    FMajorTimeUnit.Free;
  if FMinorUnit <> Nil then 
    FMinorUnit.Free;
  if FMinorTimeUnit <> Nil then 
    FMinorTimeUnit.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_DateAx.Clear;
begin
  FAssigneds := [];
  FEG_AxShared.Clear;
  if FAuto <> Nil then 
    FreeAndNil(FAuto);
  if FLblOffset <> Nil then 
    FreeAndNil(FLblOffset);
  if FBaseTimeUnit <> Nil then 
    FreeAndNil(FBaseTimeUnit);
  if FMajorUnit <> Nil then 
    FreeAndNil(FMajorUnit);
  if FMajorTimeUnit <> Nil then 
    FreeAndNil(FMajorTimeUnit);
  if FMinorUnit <> Nil then 
    FreeAndNil(FMinorUnit);
  if FMinorTimeUnit <> Nil then 
    FreeAndNil(FMinorTimeUnit);
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
end;

procedure TCT_DateAx.Assign(AItem: TCT_DateAx);
begin
end;

procedure TCT_DateAx.CopyTo(AItem: TCT_DateAx);
begin
end;

function  TCT_DateAx.Create_Auto: TCT_Boolean;
begin
  Result := TCT_Boolean.Create(FOwner);
  FAuto := Result;
end;

function  TCT_DateAx.Create_LblOffset: TCT_LblOffset;
begin
  Result := TCT_LblOffset.Create(FOwner);
  FLblOffset := Result;
end;

function  TCT_DateAx.Create_BaseTimeUnit: TCT_TimeUnit;
begin
  Result := TCT_TimeUnit.Create(FOwner);
  FBaseTimeUnit := Result;
end;

function  TCT_DateAx.Create_MajorUnit: TCT_AxisUnit;
begin
  Result := TCT_AxisUnit.Create(FOwner);
  FMajorUnit := Result;
end;

function  TCT_DateAx.Create_MajorTimeUnit: TCT_TimeUnit;
begin
  Result := TCT_TimeUnit.Create(FOwner);
  FMajorTimeUnit := Result;
end;

function  TCT_DateAx.Create_MinorUnit: TCT_AxisUnit;
begin
  Result := TCT_AxisUnit.Create(FOwner);
  FMinorUnit := Result;
end;

function  TCT_DateAx.Create_MinorTimeUnit: TCT_TimeUnit;
begin
  Result := TCT_TimeUnit.Create(FOwner);
  FMinorTimeUnit := Result;
end;

function  TCT_DateAx.Create_ExtLst: TCT_ExtensionList;
begin
  Result := TCT_ExtensionList.Create(FOwner);
  FExtLst := Result;
end;

{ TCT_DateAxXpgList }

function  TCT_DateAxXpgList.GetItems(Index: integer): TCT_DateAx;
begin
  Result := TCT_DateAx(inherited Items[Index]);
end;

function  TCT_DateAxXpgList.Add: TCT_DateAx;
begin
  Result := TCT_DateAx.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_DateAxXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_DateAxXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
    if xaElements in Items[i].Assigneds then
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
end;

procedure TCT_DateAxXpgList.Assign(AItem: TCT_DateAxXpgList);
begin
end;

procedure TCT_DateAxXpgList.CopyTo(AItem: TCT_DateAxXpgList);
begin
end;

{ TCT_SerAx }

function  TCT_SerAx.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FEG_AxShared.CheckAssigned);
  if FTickLblSkip <> Nil then 
    Inc(ElemsAssigned,FTickLblSkip.CheckAssigned);
  if FTickMarkSkip <> Nil then 
    Inc(ElemsAssigned,FTickMarkSkip.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_SerAx.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $000004F9: begin
      if FTickLblSkip = Nil then 
        FTickLblSkip := TCT_Skip.Create(FOwner);
      Result := FTickLblSkip;
    end;
    $0000056A: begin
      if FTickMarkSkip = Nil then 
        FTickMarkSkip := TCT_Skip.Create(FOwner);
      Result := FTickMarkSkip;
    end;
    $00000321: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
    begin
      Result := FEG_AxShared.HandleElement(AReader);
      if Result = Nil then 
        FOwner.Errors.Error(xemUnknownElement,AReader.QName);
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_SerAx.Write(AWriter: TXpgWriteXML);
begin
  FEG_AxShared.Write(AWriter);
  if (FTickLblSkip <> Nil) and FTickLblSkip.Assigned then 
  begin
    FTickLblSkip.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:tickLblSkip');
  end;
  if (FTickMarkSkip <> Nil) and FTickMarkSkip.Assigned then 
  begin
    FTickMarkSkip.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:tickMarkSkip');
  end;
{$ifdef XLS_WRITE_EXTLST}
  if (FExtLst <> Nil) and FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then
    begin
      AWriter.BeginTag('c:extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:extLst');
{$endif}
end;

constructor TCT_SerAx.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 4;
  FAttributeCount := 0;
  FEG_AxShared := TEG_AxShared.Create(FOwner);
end;

destructor TCT_SerAx.Destroy;
begin
  FEG_AxShared.Free;
  if FTickLblSkip <> Nil then 
    FTickLblSkip.Free;
  if FTickMarkSkip <> Nil then 
    FTickMarkSkip.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_SerAx.Clear;
begin
  FAssigneds := [];
  FEG_AxShared.Clear;
  if FTickLblSkip <> Nil then 
    FreeAndNil(FTickLblSkip);
  if FTickMarkSkip <> Nil then 
    FreeAndNil(FTickMarkSkip);
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
end;

procedure TCT_SerAx.Assign(AItem: TCT_SerAx);
begin
end;

procedure TCT_SerAx.CopyTo(AItem: TCT_SerAx);
begin
end;

function  TCT_SerAx.Create_TickLblSkip: TCT_Skip;
begin
  Result := TCT_Skip.Create(FOwner);
  FTickLblSkip := Result;
end;

function  TCT_SerAx.Create_TickMarkSkip: TCT_Skip;
begin
  Result := TCT_Skip.Create(FOwner);
  FTickMarkSkip := Result;
end;

function  TCT_SerAx.Create_ExtLst: TCT_ExtensionList;
begin
  Result := TCT_ExtensionList.Create(FOwner);
  FExtLst := Result;
end;

{ TCT_SerAxXpgList }

function  TCT_SerAxXpgList.GetItems(Index: integer): TCT_SerAx;
begin
  Result := TCT_SerAx(inherited Items[Index]);
end;

function  TCT_SerAxXpgList.Add: TCT_SerAx;
begin
  Result := TCT_SerAx.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_SerAxXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_SerAxXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
    if xaElements in Items[i].Assigneds then
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
end;

procedure TCT_SerAxXpgList.Assign(AItem: TCT_SerAxXpgList);
begin
end;

procedure TCT_SerAxXpgList.CopyTo(AItem: TCT_SerAxXpgList);
begin
end;

{ TCT_DTable }

function  TCT_DTable.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FShowHorzBorder <> Nil then 
    Inc(ElemsAssigned,FShowHorzBorder.CheckAssigned);
  if FShowVertBorder <> Nil then 
    Inc(ElemsAssigned,FShowVertBorder.CheckAssigned);
  if FShowOutline <> Nil then 
    Inc(ElemsAssigned,FShowOutline.CheckAssigned);
  if FShowKeys <> Nil then 
    Inc(ElemsAssigned,FShowKeys.CheckAssigned);
  if FSpPr <> Nil then 
    Inc(ElemsAssigned,FSpPr.CheckAssigned);
  if FTxPr <> Nil then 
    Inc(ElemsAssigned,FTxPr.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_DTable.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $0000065F: begin
      if FShowHorzBorder = Nil then 
        FShowHorzBorder := TCT_Boolean.Create(FOwner);
      Result := FShowHorzBorder;
    end;
    $0000065D: begin
      if FShowVertBorder = Nil then 
        FShowVertBorder := TCT_Boolean.Create(FOwner);
      Result := FShowVertBorder;
    end;
    $0000053E: begin
      if FShowOutline = Nil then 
        FShowOutline := TCT_Boolean.Create(FOwner);
      Result := FShowOutline;
    end;
    $000003FA: begin
      if FShowKeys = Nil then 
        FShowKeys := TCT_Boolean.Create(FOwner);
      Result := FShowKeys;
    end;
    $00000242: begin
      if FSpPr = Nil then 
        FSpPr := TCT_ShapeProperties.Create(FOwner);
      Result := FSpPr;
    end;
    $0000024B: begin
      if FTxPr = Nil then 
        FTxPr := TCT_TextBody.Create(FOwner);
      Result := FTxPr;
    end;
    $00000321: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_DTable.Write(AWriter: TXpgWriteXML);
begin
  if (FShowHorzBorder <> Nil) and FShowHorzBorder.Assigned then 
  begin
    FShowHorzBorder.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:showHorzBorder');
  end;
  if (FShowVertBorder <> Nil) and FShowVertBorder.Assigned then 
  begin
    FShowVertBorder.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:showVertBorder');
  end;
  if (FShowOutline <> Nil) and FShowOutline.Assigned then 
  begin
    FShowOutline.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:showOutline');
  end;
  if (FShowKeys <> Nil) and FShowKeys.Assigned then 
  begin
    FShowKeys.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:showKeys');
  end;
  if (FSpPr <> Nil) and FSpPr.Assigned then 
  begin
    FSpPr.WriteAttributes(AWriter);
    if xaElements in FSpPr.Assigneds then
    begin
      AWriter.BeginTag('c:spPr');
      FSpPr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:spPr');
  end;
  if (FTxPr <> Nil) and FTxPr.Assigned then 
    if xaElements in FTxPr.Assigneds then
    begin
      AWriter.BeginTag('c:txPr');
      FTxPr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:txPr');
{$ifdef XLS_WRITE_EXTLST}
  if (FExtLst <> Nil) and FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then
    begin
      AWriter.BeginTag('c:extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:extLst');
{$endif}
end;

constructor TCT_DTable.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 7;
  FAttributeCount := 0;
end;

destructor TCT_DTable.Destroy;
begin
  if FShowHorzBorder <> Nil then 
    FShowHorzBorder.Free;
  if FShowVertBorder <> Nil then 
    FShowVertBorder.Free;
  if FShowOutline <> Nil then 
    FShowOutline.Free;
  if FShowKeys <> Nil then 
    FShowKeys.Free;
  if FSpPr <> Nil then 
    FSpPr.Free;
  if FTxPr <> Nil then 
    FTxPr.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_DTable.Clear;
begin
  FAssigneds := [];
  if FShowHorzBorder <> Nil then 
    FreeAndNil(FShowHorzBorder);
  if FShowVertBorder <> Nil then 
    FreeAndNil(FShowVertBorder);
  if FShowOutline <> Nil then 
    FreeAndNil(FShowOutline);
  if FShowKeys <> Nil then 
    FreeAndNil(FShowKeys);
  if FSpPr <> Nil then 
    FreeAndNil(FSpPr);
  if FTxPr <> Nil then 
    FreeAndNil(FTxPr);
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
end;

procedure TCT_DTable.Assign(AItem: TCT_DTable);
begin
end;

procedure TCT_DTable.CopyTo(AItem: TCT_DTable);
begin
end;

function  TCT_DTable.Create_ShowHorzBorder: TCT_Boolean;
begin
  Result := TCT_Boolean.Create(FOwner);
  FShowHorzBorder := Result;
end;

function  TCT_DTable.Create_ShowVertBorder: TCT_Boolean;
begin
  Result := TCT_Boolean.Create(FOwner);
  FShowVertBorder := Result;
end;

function  TCT_DTable.Create_ShowOutline: TCT_Boolean;
begin
  Result := TCT_Boolean.Create(FOwner);
  FShowOutline := Result;
end;

function  TCT_DTable.Create_ShowKeys: TCT_Boolean;
begin
  Result := TCT_Boolean.Create(FOwner);
  FShowKeys := Result;
end;

function  TCT_DTable.Create_SpPr: TCT_ShapeProperties;
begin
  Result := TCT_ShapeProperties.Create(FOwner);
  FSpPr := Result;
end;

function  TCT_DTable.Create_TxPr: TCT_TextBody;
begin
  Result := TCT_TextBody.Create(FOwner);
  FTxPr := Result;
end;

function  TCT_DTable.Create_ExtLst: TCT_ExtensionList;
begin
  Result := TCT_ExtensionList.Create(FOwner);
  FExtLst := Result;
end;

{ TCT_LegendPos }

function  TCT_LegendPos.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FVal <> stlpR then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_LegendPos.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_LegendPos.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FVal <> stlpR then 
    AWriter.AddAttribute('val',StrTST_LegendPos[Integer(FVal)]);
end;

procedure TCT_LegendPos.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then 
    FVal := TST_LegendPos(StrToEnum('stlp' + AAttributes.Values[0]))
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_LegendPos.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := stlpR;
end;

destructor TCT_LegendPos.Destroy;
begin
end;

procedure TCT_LegendPos.Clear;
begin
  FAssigneds := [];
  FVal := stlpR;
end;

procedure TCT_LegendPos.Assign(AItem: TCT_LegendPos);
begin
end;

procedure TCT_LegendPos.CopyTo(AItem: TCT_LegendPos);
begin
end;

{ TCT_LegendEntry }

function  TCT_LegendEntry.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FIdx <> Nil then 
    Inc(ElemsAssigned,FIdx.CheckAssigned);
  if FDelete <> Nil then 
    Inc(ElemsAssigned,FDelete.CheckAssigned);
  Inc(ElemsAssigned,FEG_LegendEntryData.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_LegendEntry.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $000001E2: begin
      if FIdx = Nil then 
        FIdx := TCT_UnsignedInt.Create(FOwner);
      Result := FIdx;
    end;
    $00000310: begin
      if FDelete = Nil then 
        FDelete := TCT_Boolean.Create(FOwner);
      Result := FDelete;
    end;
    $00000321: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
    begin
      Result := FEG_LegendEntryData.HandleElement(AReader);
      if Result = Nil then 
        FOwner.Errors.Error(xemUnknownElement,AReader.QName);
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_LegendEntry.Write(AWriter: TXpgWriteXML);
begin
  if (FIdx <> Nil) and FIdx.Assigned then 
  begin
    FIdx.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:idx');
  end
  else 
    AWriter.SimpleTag('c:idx');
  if (FDelete <> Nil) and FDelete.Assigned then 
  begin
    FDelete.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:delete');
  end;
  FEG_LegendEntryData.Write(AWriter);
{$ifdef XLS_WRITE_EXTLST}
  if (FExtLst <> Nil) and FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then
    begin
      AWriter.BeginTag('c:extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:extLst');
{$endif}
end;

constructor TCT_LegendEntry.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 4;
  FAttributeCount := 0;
  FEG_LegendEntryData := TEG_LegendEntryData.Create(FOwner);
end;

destructor TCT_LegendEntry.Destroy;
begin
  if FIdx <> Nil then 
    FIdx.Free;
  if FDelete <> Nil then 
    FDelete.Free;
  FEG_LegendEntryData.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_LegendEntry.Clear;
begin
  FAssigneds := [];
  if FIdx <> Nil then 
    FreeAndNil(FIdx);
  if FDelete <> Nil then 
    FreeAndNil(FDelete);
  FEG_LegendEntryData.Clear;
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
end;

procedure TCT_LegendEntry.Assign(AItem: TCT_LegendEntry);
begin
end;

procedure TCT_LegendEntry.CopyTo(AItem: TCT_LegendEntry);
begin
end;

function  TCT_LegendEntry.Create_Idx: TCT_UnsignedInt;
begin
  Result := TCT_UnsignedInt.Create(FOwner);
  FIdx := Result;
end;

function  TCT_LegendEntry.Create_Delete: TCT_Boolean;
begin
  Result := TCT_Boolean.Create(FOwner);
  FDelete := Result;
end;

function  TCT_LegendEntry.Create_ExtLst: TCT_ExtensionList;
begin
  Result := TCT_ExtensionList.Create(FOwner);
  FExtLst := Result;
end;

{ TCT_LegendEntryXpgList }

function  TCT_LegendEntryXpgList.GetItems(Index: integer): TCT_LegendEntry;
begin
  Result := TCT_LegendEntry(inherited Items[Index]);
end;

function  TCT_LegendEntryXpgList.Add: TCT_LegendEntry;
begin
  Result := TCT_LegendEntry.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_LegendEntryXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_LegendEntryXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
    if xaElements in Items[i].Assigneds then
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
end;

procedure TCT_LegendEntryXpgList.Assign(AItem: TCT_LegendEntryXpgList);
begin
end;

procedure TCT_LegendEntryXpgList.CopyTo(AItem: TCT_LegendEntryXpgList);
begin
end;

{ TCT_DefaultShapeDefinition }

function  TCT_DefaultShapeDefinition.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FA_SpPr <> Nil then 
    Inc(ElemsAssigned,FA_SpPr.CheckAssigned);
  if FA_BodyPr <> Nil then 
    Inc(ElemsAssigned,FA_BodyPr.CheckAssigned);
  if FA_LstStyle <> Nil then 
    Inc(ElemsAssigned,FA_LstStyle.CheckAssigned);
  if FA_Style <> Nil then 
    Inc(ElemsAssigned,FA_Style.CheckAssigned);
  if FA_ExtLst <> Nil then 
    Inc(ElemsAssigned,FA_ExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_DefaultShapeDefinition.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $00000240: begin
      if FA_SpPr = Nil then 
        FA_SpPr := TCT_ShapeProperties.Create(FOwner);
      Result := FA_SpPr;
    end;
    $0000030B: begin
      if FA_BodyPr = Nil then 
        FA_BodyPr := TCT_TextBodyProperties.Create(FOwner);
      Result := FA_BodyPr;
    end;
    $000003FF: begin
      if FA_LstStyle = Nil then 
        FA_LstStyle := TCT_TextListStyle.Create(FOwner);
      Result := FA_LstStyle;
    end;
    $000002CC: begin
      if FA_Style = Nil then 
        FA_Style := TCT_ShapeStyle.Create(FOwner);
      Result := FA_Style;
    end;
    $0000031F: begin
      if FA_ExtLst = Nil then 
        FA_ExtLst := TCT_OfficeArtExtensionList.Create(FOwner);
      Result := FA_ExtLst;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_DefaultShapeDefinition.Write(AWriter: TXpgWriteXML);
begin
  if (FA_SpPr <> Nil) and FA_SpPr.Assigned then
  begin
    FA_SpPr.WriteAttributes(AWriter);
    if xaElements in FA_SpPr.Assigneds then
    begin
      AWriter.BeginTag('a:spPr');
      FA_SpPr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:spPr');
  end
  else 
    AWriter.SimpleTag('a:spPr');
  if (FA_BodyPr <> Nil) and FA_BodyPr.Assigned then 
  begin
    FA_BodyPr.WriteAttributes(AWriter);
    if xaElements in FA_BodyPr.Assigneds then
    begin
      AWriter.BeginTag('a:bodyPr');
      FA_BodyPr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:bodyPr');
  end
  else 
    AWriter.SimpleTag('a:bodyPr');

  if (FA_LstStyle <> Nil) and FA_LstStyle.Assigned then
    if xaElements in FA_LstStyle.Assigneds then
    begin
      AWriter.BeginTag('a:lstStyle');
      FA_LstStyle.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('a:lstStyle')
  else 
    AWriter.SimpleTag('a:lstStyle');
  if (FA_Style <> Nil) and FA_Style.Assigned then 
    if xaElements in FA_Style.Assigneds then
    begin
      AWriter.BeginTag('a:style');
      FA_Style.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:style');
  if (FA_ExtLst <> Nil) and FA_ExtLst.Assigned then 
    if xaElements in FA_ExtLst.Assigneds then
    begin
      AWriter.BeginTag('a:extLst');
      FA_ExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:extLst');
end;

constructor TCT_DefaultShapeDefinition.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 5;
  FAttributeCount := 0;
end;

destructor TCT_DefaultShapeDefinition.Destroy;
begin
  if FA_SpPr <> Nil then 
    FA_SpPr.Free;
  if FA_BodyPr <> Nil then 
    FA_BodyPr.Free;
  if FA_LstStyle <> Nil then 
    FA_LstStyle.Free;
  if FA_Style <> Nil then 
    FA_Style.Free;
  if FA_ExtLst <> Nil then 
    FA_ExtLst.Free;
end;

procedure TCT_DefaultShapeDefinition.Clear;
begin
  FAssigneds := [];
  if FA_SpPr <> Nil then 
    FreeAndNil(FA_SpPr);
  if FA_BodyPr <> Nil then 
    FreeAndNil(FA_BodyPr);
  if FA_LstStyle <> Nil then 
    FreeAndNil(FA_LstStyle);
  if FA_Style <> Nil then 
    FreeAndNil(FA_Style);
  if FA_ExtLst <> Nil then 
    FreeAndNil(FA_ExtLst);
end;

procedure TCT_DefaultShapeDefinition.Assign(AItem: TCT_DefaultShapeDefinition);
begin
end;

procedure TCT_DefaultShapeDefinition.CopyTo(AItem: TCT_DefaultShapeDefinition);
begin
end;

function  TCT_DefaultShapeDefinition.Create_A_SpPr: TCT_ShapeProperties;
begin
  Result := TCT_ShapeProperties.Create(FOwner);
  FA_SpPr := Result;
end;

function  TCT_DefaultShapeDefinition.Create_A_BodyPr: TCT_TextBodyProperties;
begin
  Result := TCT_TextBodyProperties.Create(FOwner);
  FA_BodyPr := Result;
end;

function  TCT_DefaultShapeDefinition.Create_A_LstStyle: TCT_TextListStyle;
begin
  Result := TCT_TextListStyle.Create(FOwner);
  FA_LstStyle := Result;
end;

function  TCT_DefaultShapeDefinition.Create_A_Style: TCT_ShapeStyle;
begin
  Result := TCT_ShapeStyle.Create(FOwner);
  FA_Style := Result;
end;

function  TCT_DefaultShapeDefinition.Create_A_ExtLst: TCT_OfficeArtExtensionList;
begin
  Result := TCT_OfficeArtExtensionList.Create(FOwner);
  FA_ExtLst := Result;
end;

{ TCT_ColorSchemeAndMapping }

function  TCT_ColorSchemeAndMapping.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FA_ClrScheme <> Nil then 
    Inc(ElemsAssigned,FA_ClrScheme.CheckAssigned);
  if FA_ClrMap <> Nil then 
    Inc(ElemsAssigned,FA_ClrMap.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_ColorSchemeAndMapping.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $00000431: begin
      if FA_ClrScheme = Nil then 
        FA_ClrScheme := TCT_ColorScheme.Create(FOwner);
      Result := FA_ClrScheme;
    end;
    $000002FA: begin
      if FA_ClrMap = Nil then 
        FA_ClrMap := TCT_ColorMapping.Create(FOwner);
      Result := FA_ClrMap;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_ColorSchemeAndMapping.Write(AWriter: TXpgWriteXML);
begin
  if (FA_ClrScheme <> Nil) and FA_ClrScheme.Assigned then 
  begin
    FA_ClrScheme.WriteAttributes(AWriter);
    if xaElements in FA_ClrScheme.Assigneds then
    begin
      AWriter.BeginTag('a:clrScheme');
      FA_ClrScheme.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:clrScheme');
  end
  else 
    AWriter.SimpleTag('a:clrScheme');
  if (FA_ClrMap <> Nil) and FA_ClrMap.Assigned then 
  begin
    FA_ClrMap.WriteAttributes(AWriter);
    if xaElements in FA_ClrMap.Assigneds then
    begin
      AWriter.BeginTag('a:clrMap');
      FA_ClrMap.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:clrMap');
  end;
end;

constructor TCT_ColorSchemeAndMapping.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 0;
end;

destructor TCT_ColorSchemeAndMapping.Destroy;
begin
  if FA_ClrScheme <> Nil then 
    FA_ClrScheme.Free;
  if FA_ClrMap <> Nil then 
    FA_ClrMap.Free;
end;

procedure TCT_ColorSchemeAndMapping.Clear;
begin
  FAssigneds := [];
  if FA_ClrScheme <> Nil then 
    FreeAndNil(FA_ClrScheme);
  if FA_ClrMap <> Nil then 
    FreeAndNil(FA_ClrMap);
end;

procedure TCT_ColorSchemeAndMapping.Assign(AItem: TCT_ColorSchemeAndMapping);
begin
end;

procedure TCT_ColorSchemeAndMapping.CopyTo(AItem: TCT_ColorSchemeAndMapping);
begin
end;

function  TCT_ColorSchemeAndMapping.Create_A_ClrScheme: TCT_ColorScheme;
begin
  Result := TCT_ColorScheme.Create(FOwner);
  FA_ClrScheme := Result;
end;

function  TCT_ColorSchemeAndMapping.Create_A_ClrMap: TCT_ColorMapping;
begin
  Result := TCT_ColorMapping.Create(FOwner);
  FA_ClrMap := Result;
end;

{ TCT_ColorSchemeAndMappingXpgList }

function  TCT_ColorSchemeAndMappingXpgList.GetItems(Index: integer): TCT_ColorSchemeAndMapping;
begin
  Result := TCT_ColorSchemeAndMapping(inherited Items[Index]);
end;

function  TCT_ColorSchemeAndMappingXpgList.Add: TCT_ColorSchemeAndMapping;
begin
  Result := TCT_ColorSchemeAndMapping.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_ColorSchemeAndMappingXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_ColorSchemeAndMappingXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
    if xaElements in Items[i].Assigneds then
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
end;

procedure TCT_ColorSchemeAndMappingXpgList.Assign(AItem: TCT_ColorSchemeAndMappingXpgList);
begin
end;

procedure TCT_ColorSchemeAndMappingXpgList.CopyTo(AItem: TCT_ColorSchemeAndMappingXpgList);
begin
end;

{ TCT_PivotFmts }

function  TCT_PivotFmts.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FPivotFmtXpgList <> Nil then 
    Inc(ElemsAssigned,FPivotFmtXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_PivotFmts.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'c:pivotFmt' then 
  begin
    if FPivotFmtXpgList = Nil then 
      FPivotFmtXpgList := TCT_PivotFmtXpgList.Create(FOwner);
    Result := FPivotFmtXpgList.Add;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_PivotFmts.Write(AWriter: TXpgWriteXML);
begin
  if FPivotFmtXpgList <> Nil then 
    FPivotFmtXpgList.Write(AWriter,'c:pivotFmt');
end;

constructor TCT_PivotFmts.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
end;

destructor TCT_PivotFmts.Destroy;
begin
  if FPivotFmtXpgList <> Nil then 
    FPivotFmtXpgList.Free;
end;

procedure TCT_PivotFmts.Clear;
begin
  FAssigneds := [];
  if FPivotFmtXpgList <> Nil then 
    FreeAndNil(FPivotFmtXpgList);
end;

procedure TCT_PivotFmts.Assign(AItem: TCT_PivotFmts);
begin
end;

procedure TCT_PivotFmts.CopyTo(AItem: TCT_PivotFmts);
begin
end;

function  TCT_PivotFmts.Create_PivotFmtXpgList: TCT_PivotFmtXpgList;
begin
  Result := TCT_PivotFmtXpgList.Create(FOwner);
  FPivotFmtXpgList := Result;
end;

{ TCT_View3D }

function  TCT_View3D.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FRotX <> Nil then 
    Inc(ElemsAssigned,FRotX.CheckAssigned);
  if FHPercent <> Nil then 
    Inc(ElemsAssigned,FHPercent.CheckAssigned);
  if FRotY <> Nil then 
    Inc(ElemsAssigned,FRotY.CheckAssigned);
  if FDepthPercent <> Nil then 
    Inc(ElemsAssigned,FDepthPercent.CheckAssigned);
  if FRAngAx <> Nil then 
    Inc(ElemsAssigned,FRAngAx.CheckAssigned);
  if FPerspective <> Nil then 
    Inc(ElemsAssigned,FPerspective.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_View3D.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $0000024A: begin
      if FRotX = Nil then 
        FRotX := TCT_RotX.Create(FOwner);
      Result := FRotX;
    end;
    $000003D6: begin
      if FHPercent = Nil then 
        FHPercent := TCT_HPercent.Create(FOwner);
      Result := FHPercent;
    end;
    $0000024B: begin
      if FRotY = Nil then 
        FRotY := TCT_RotY.Create(FOwner);
      Result := FRotY;
    end;
    $00000583: begin
      if FDepthPercent = Nil then 
        FDepthPercent := TCT_DepthPercent.Create(FOwner);
      Result := FDepthPercent;
    end;
    $000002DE: begin
      if FRAngAx = Nil then 
        FRAngAx := TCT_Boolean.Create(FOwner);
      Result := FRAngAx;
    end;
    $00000547: begin
      if FPerspective = Nil then 
        FPerspective := TCT_Perspective.Create(FOwner);
      Result := FPerspective;
    end;
    $00000321: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_View3D.Write(AWriter: TXpgWriteXML);
begin
  if (FRotX <> Nil) and FRotX.Assigned then 
  begin
    FRotX.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:rotX');
  end;
  if (FHPercent <> Nil) and FHPercent.Assigned then 
  begin
    FHPercent.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:hPercent');
  end;
  if (FRotY <> Nil) and FRotY.Assigned then 
  begin
    FRotY.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:rotY');
  end;
  if (FDepthPercent <> Nil) and FDepthPercent.Assigned then 
  begin
    FDepthPercent.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:depthPercent');
  end;
  if (FRAngAx <> Nil) and FRAngAx.Assigned then 
  begin
    FRAngAx.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:rAngAx');
  end;
  if (FPerspective <> Nil) and FPerspective.Assigned then 
  begin
    FPerspective.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:perspective');
  end;
{$ifdef XLS_WRITE_EXTLST}
  if (FExtLst <> Nil) and FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then
    begin
      AWriter.BeginTag('c:extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:extLst');
{$endif}
end;

constructor TCT_View3D.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 7;
  FAttributeCount := 0;
end;

destructor TCT_View3D.Destroy;
begin
  if FRotX <> Nil then 
    FRotX.Free;
  if FHPercent <> Nil then 
    FHPercent.Free;
  if FRotY <> Nil then 
    FRotY.Free;
  if FDepthPercent <> Nil then 
    FDepthPercent.Free;
  if FRAngAx <> Nil then 
    FRAngAx.Free;
  if FPerspective <> Nil then 
    FPerspective.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_View3D.Clear;
begin
  FAssigneds := [];
  if FRotX <> Nil then 
    FreeAndNil(FRotX);
  if FHPercent <> Nil then 
    FreeAndNil(FHPercent);
  if FRotY <> Nil then 
    FreeAndNil(FRotY);
  if FDepthPercent <> Nil then 
    FreeAndNil(FDepthPercent);
  if FRAngAx <> Nil then 
    FreeAndNil(FRAngAx);
  if FPerspective <> Nil then 
    FreeAndNil(FPerspective);
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
end;

procedure TCT_View3D.Assign(AItem: TCT_View3D);
begin
end;

procedure TCT_View3D.CopyTo(AItem: TCT_View3D);
begin
end;

function  TCT_View3D.Create_RotX: TCT_RotX;
begin
  Result := TCT_RotX.Create(FOwner);
  FRotX := Result;
end;

function  TCT_View3D.Create_HPercent: TCT_HPercent;
begin
  Result := TCT_HPercent.Create(FOwner);
  FHPercent := Result;
end;

function  TCT_View3D.Create_RotY: TCT_RotY;
begin
  Result := TCT_RotY.Create(FOwner);
  FRotY := Result;
end;

function  TCT_View3D.Create_DepthPercent: TCT_DepthPercent;
begin
  Result := TCT_DepthPercent.Create(FOwner);
  FDepthPercent := Result;
end;

function  TCT_View3D.Create_RAngAx: TCT_Boolean;
begin
  Result := TCT_Boolean.Create(FOwner);
  FRAngAx := Result;
end;

function  TCT_View3D.Create_Perspective: TCT_Perspective;
begin
  Result := TCT_Perspective.Create(FOwner);
  FPerspective := Result;
end;

function  TCT_View3D.Create_ExtLst: TCT_ExtensionList;
begin
  Result := TCT_ExtensionList.Create(FOwner);
  FExtLst := Result;
end;

{ TCT_Surface }

function  TCT_Surface.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FThickness <> Nil then 
    Inc(ElemsAssigned,FThickness.CheckAssigned);
  if FSpPr <> Nil then 
    Inc(ElemsAssigned,FSpPr.CheckAssigned);
  if FPictureOptions <> Nil then 
    Inc(ElemsAssigned,FPictureOptions.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_Surface.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $00000469: begin
      if FThickness = Nil then 
        FThickness := TCT_UnsignedInt.Create(FOwner);
      Result := FThickness;
    end;
    $00000242: begin
      if FSpPr = Nil then 
        FSpPr := TCT_ShapeProperties.Create(FOwner);
      Result := FSpPr;
    end;
    $00000685: begin
      if FPictureOptions = Nil then 
        FPictureOptions := TCT_PictureOptions.Create(FOwner);
      Result := FPictureOptions;
    end;
    $00000321: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_Surface.Write(AWriter: TXpgWriteXML);
begin
  if (FThickness <> Nil) and FThickness.Assigned then 
  begin
    FThickness.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:thickness');
  end;
  if (FSpPr <> Nil) and FSpPr.Assigned then 
  begin
    FSpPr.WriteAttributes(AWriter);
    if xaElements in FSpPr.Assigneds then
    begin
      AWriter.BeginTag('c:spPr');
      FSpPr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:spPr');
  end;
  if (FPictureOptions <> Nil) and FPictureOptions.Assigned then 
    if xaElements in FPictureOptions.Assigneds then
    begin
      AWriter.BeginTag('c:pictureOptions');
      FPictureOptions.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:pictureOptions');
{$ifdef XLS_WRITE_EXTLST}
  if (FExtLst <> Nil) and FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then
    begin
      AWriter.BeginTag('c:extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:extLst');
{$endif}
end;

constructor TCT_Surface.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 4;
  FAttributeCount := 0;
end;

destructor TCT_Surface.Destroy;
begin
  if FThickness <> Nil then 
    FThickness.Free;
  if FSpPr <> Nil then 
    FSpPr.Free;
  if FPictureOptions <> Nil then 
    FPictureOptions.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_Surface.Clear;
begin
  FAssigneds := [];
  if FThickness <> Nil then 
    FreeAndNil(FThickness);
  if FSpPr <> Nil then 
    FreeAndNil(FSpPr);
  if FPictureOptions <> Nil then 
    FreeAndNil(FPictureOptions);
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
end;

procedure TCT_Surface.Assign(AItem: TCT_Surface);
begin
end;

procedure TCT_Surface.CopyTo(AItem: TCT_Surface);
begin
end;

function  TCT_Surface.Create_Thickness: TCT_UnsignedInt;
begin
  Result := TCT_UnsignedInt.Create(FOwner);
  FThickness := Result;
end;

function  TCT_Surface.Create_SpPr: TCT_ShapeProperties;
begin
  Result := TCT_ShapeProperties.Create(FOwner);
  FSpPr := Result;
end;

function  TCT_Surface.Create_PictureOptions: TCT_PictureOptions;
begin
  Result := TCT_PictureOptions.Create(FOwner);
  FPictureOptions := Result;
end;

function  TCT_Surface.Create_ExtLst: TCT_ExtensionList;
begin
  Result := TCT_ExtensionList.Create(FOwner);
  FExtLst := Result;
end;

{ TCT_PlotArea }

function  TCT_PlotArea.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FLayout <> Nil then 
    Inc(ElemsAssigned,FLayout.CheckAssigned);
  if FAreaChartXpgList <> Nil then 
    Inc(ElemsAssigned,FAreaChartXpgList.CheckAssigned);
  if FArea3DChartXpgList <> Nil then 
    Inc(ElemsAssigned,FArea3DChartXpgList.CheckAssigned);
  if FLineChartXpgList <> Nil then 
    Inc(ElemsAssigned,FLineChartXpgList.CheckAssigned);
  if FLine3DChartXpgList <> Nil then 
    Inc(ElemsAssigned,FLine3DChartXpgList.CheckAssigned);
  if FStockChartXpgList <> Nil then 
    Inc(ElemsAssigned,FStockChartXpgList.CheckAssigned);
  if FRadarChartXpgList <> Nil then 
    Inc(ElemsAssigned,FRadarChartXpgList.CheckAssigned);
  if FScatterChartXpgList <> Nil then 
    Inc(ElemsAssigned,FScatterChartXpgList.CheckAssigned);
  if FPieChartXpgList <> Nil then 
    Inc(ElemsAssigned,FPieChartXpgList.CheckAssigned);
  if FPie3DChartXpgList <> Nil then 
    Inc(ElemsAssigned,FPie3DChartXpgList.CheckAssigned);
  if FDoughnutChartXpgList <> Nil then 
    Inc(ElemsAssigned,FDoughnutChartXpgList.CheckAssigned);
  if FBarChartXpgList <> Nil then 
    Inc(ElemsAssigned,FBarChartXpgList.CheckAssigned);
  if FBar3DChartXpgList <> Nil then 
    Inc(ElemsAssigned,FBar3DChartXpgList.CheckAssigned);
  if FOfPieChartXpgList <> Nil then 
    Inc(ElemsAssigned,FOfPieChartXpgList.CheckAssigned);
  if FSurfaceChartXpgList <> Nil then 
    Inc(ElemsAssigned,FSurfaceChartXpgList.CheckAssigned);
  if FSurface3DChartXpgList <> Nil then 
    Inc(ElemsAssigned,FSurface3DChartXpgList.CheckAssigned);
  if FBubbleChartXpgList <> Nil then 
    Inc(ElemsAssigned,FBubbleChartXpgList.CheckAssigned);
  if FValAxXpgList <> Nil then 
    Inc(ElemsAssigned,FValAxXpgList.CheckAssigned);
  if FCatAxXpgList <> Nil then 
    Inc(ElemsAssigned,FCatAxXpgList.CheckAssigned);
  if FDateAxXpgList <> Nil then 
    Inc(ElemsAssigned,FDateAxXpgList.CheckAssigned);
  if FSerAxXpgList <> Nil then 
    Inc(ElemsAssigned,FSerAxXpgList.CheckAssigned);
  if FDTable <> Nil then 
    Inc(ElemsAssigned,FDTable.CheckAssigned);
  if FSpPr <> Nil then 
    Inc(ElemsAssigned,FSpPr.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_PlotArea.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $0000033B: begin
      if FLayout = Nil then 
        FLayout := TCT_Layout.Create(FOwner);
      Result := FLayout;
    end;
    $00000428: begin
      if FAreaChartXpgList = Nil then 
        FAreaChartXpgList := TCT_AreaChartXpgList.Create(FOwner);
      Result := FAreaChartXpgList.Add;
    end;
    $0000049F: begin
      if FArea3DChartXpgList = Nil then 
        FArea3DChartXpgList := TCT_Area3DChartXpgList.Create(FOwner);
      Result := FArea3DChartXpgList.Add;
    end;
    $00000437: begin
      if FLineChartXpgList = Nil then 
        FLineChartXpgList := TCT_LineChartXpgList.Create(FOwner);
      Result := FLineChartXpgList.Add;
    end;
    $000004AE: begin
      if FLine3DChartXpgList = Nil then 
        FLine3DChartXpgList := TCT_Line3DChartXpgList.Create(FOwner);
      Result := FLine3DChartXpgList.Add;
    end;
    $000004B3: begin
      if FStockChartXpgList = Nil then 
        FStockChartXpgList := TCT_StockChartXpgList.Create(FOwner);
      Result := FStockChartXpgList.Add;
    end;
    $00000499: begin
      if FRadarChartXpgList = Nil then 
        FRadarChartXpgList := TCT_RadarChartXpgList.Create(FOwner);
      Result := FRadarChartXpgList.Add;
    end;
    $00000585: begin
      if FScatterChartXpgList = Nil then 
        FScatterChartXpgList := TCT_ScatterChartXpgList.Create(FOwner);
      Result := FScatterChartXpgList.Add;
    end;
    $000003CD: begin
      if FPieChartXpgList = Nil then 
        FPieChartXpgList := TCT_PieChartXpgList.Create(FOwner);
      Result := FPieChartXpgList.Add;
    end;
    $00000444: begin
      if FPie3DChartXpgList = Nil then 
        FPie3DChartXpgList := TCT_Pie3DChartXpgList.Create(FOwner);
      Result := FPie3DChartXpgList.Add;
    end;
    $000005FD: begin
      if FDoughnutChartXpgList = Nil then 
        FDoughnutChartXpgList := TCT_DoughnutChartXpgList.Create(FOwner);
      Result := FDoughnutChartXpgList.Add;
    end;
    $000003C4: begin
      if FBarChartXpgList = Nil then 
        FBarChartXpgList := TCT_BarChartXpgList.Create(FOwner);
      Result := FBarChartXpgList.Add;
    end;
    $0000043B: begin
      if FBar3DChartXpgList = Nil then 
        FBar3DChartXpgList := TCT_Bar3DChartXpgList.Create(FOwner);
      Result := FBar3DChartXpgList.Add;
    end;
    $00000482: begin
      if FOfPieChartXpgList = Nil then 
        FOfPieChartXpgList := TCT_OfPieChartXpgList.Create(FOwner);
      Result := FOfPieChartXpgList.Add;
    end;
    $00000578: begin
      if FSurfaceChartXpgList = Nil then 
        FSurfaceChartXpgList := TCT_SurfaceChartXpgList.Create(FOwner);
      Result := FSurfaceChartXpgList.Add;
    end;
    $000005EF: begin
      if FSurface3DChartXpgList = Nil then 
        FSurface3DChartXpgList := TCT_Surface3DChartXpgList.Create(FOwner);
      Result := FSurface3DChartXpgList.Add;
    end;
    $000004FB: begin
      if FBubbleChartXpgList = Nil then 
        FBubbleChartXpgList := TCT_BubbleChartXpgList.Create(FOwner);
      Result := FBubbleChartXpgList.Add;
    end;
    $00000299: begin
      if FValAxXpgList = Nil then 
        FValAxXpgList := TCT_ValAxXpgList.Create(FOwner);
      Result := FValAxXpgList.Add;
    end;
    $0000028E: begin
      if FCatAxXpgList = Nil then 
        FCatAxXpgList := TCT_CatAxXpgList.Create(FOwner);
      Result := FCatAxXpgList.Add;
    end;
    $000002F4: begin
      if FDateAxXpgList = Nil then 
        FDateAxXpgList := TCT_DateAxXpgList.Create(FOwner);
      Result := FDateAxXpgList.Add;
    end;
    $000002A0: begin
      if FSerAxXpgList = Nil then 
        FSerAxXpgList := TCT_SerAxXpgList.Create(FOwner);
      Result := FSerAxXpgList.Add;
    end;
    $000002E9: begin
      if FDTable = Nil then 
        FDTable := TCT_DTable.Create(FOwner);
      Result := FDTable;
    end;
    $00000242: begin
      if FSpPr = Nil then 
        FSpPr := TCT_ShapeProperties.Create(FOwner);
      Result := FSpPr;
    end;
    $00000321: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_PlotArea.Write(AWriter: TXpgWriteXML);
begin
  if (FLayout <> Nil) and FLayout.Assigned and (xaElements in FLayout.Assigneds) then begin
    AWriter.BeginTag('c:layout');
    FLayout.Write(AWriter);
    AWriter.EndTag;
  end
  else
    AWriter.SimpleTag('c:layout');
  if FAreaChartXpgList <> Nil then
    FAreaChartXpgList.Write(AWriter,'c:areaChart');
  if FArea3DChartXpgList <> Nil then 
    FArea3DChartXpgList.Write(AWriter,'c:area3DChart');
  if FLineChartXpgList <> Nil then 
    FLineChartXpgList.Write(AWriter,'c:lineChart');
  if FLine3DChartXpgList <> Nil then 
    FLine3DChartXpgList.Write(AWriter,'c:line3DChart');
  if FStockChartXpgList <> Nil then 
    FStockChartXpgList.Write(AWriter,'c:stockChart');
  if FRadarChartXpgList <> Nil then 
    FRadarChartXpgList.Write(AWriter,'c:radarChart');
  if FScatterChartXpgList <> Nil then 
    FScatterChartXpgList.Write(AWriter,'c:scatterChart');
  if FPieChartXpgList <> Nil then 
    FPieChartXpgList.Write(AWriter,'c:pieChart');
  if FPie3DChartXpgList <> Nil then 
    FPie3DChartXpgList.Write(AWriter,'c:pie3DChart');
  if FDoughnutChartXpgList <> Nil then 
    FDoughnutChartXpgList.Write(AWriter,'c:doughnutChart');
  if FBarChartXpgList <> Nil then 
    FBarChartXpgList.Write(AWriter,'c:barChart');
  if FBar3DChartXpgList <> Nil then 
    FBar3DChartXpgList.Write(AWriter,'c:bar3DChart');
  if FOfPieChartXpgList <> Nil then 
    FOfPieChartXpgList.Write(AWriter,'c:ofPieChart');
  if FSurfaceChartXpgList <> Nil then 
    FSurfaceChartXpgList.Write(AWriter,'c:surfaceChart');
  if FSurface3DChartXpgList <> Nil then 
    FSurface3DChartXpgList.Write(AWriter,'c:surface3DChart');
  if FBubbleChartXpgList <> Nil then 
    FBubbleChartXpgList.Write(AWriter,'c:bubbleChart');
  if FCatAxXpgList <> Nil then
    FCatAxXpgList.Write(AWriter,'c:catAx');
  if FDateAxXpgList <> Nil then
    FDateAxXpgList.Write(AWriter,'c:dateAx');
  if FValAxXpgList <> Nil then
    FValAxXpgList.Write(AWriter,'c:valAx');
  if FSerAxXpgList <> Nil then
    FSerAxXpgList.Write(AWriter,'c:serAx');
  if (FDTable <> Nil) and FDTable.Assigned then 
    if xaElements in FDTable.Assigneds then
    begin
      AWriter.BeginTag('c:dTable');
      FDTable.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:dTable');
  if (FSpPr <> Nil) and FSpPr.Assigned then 
  begin
    FSpPr.WriteAttributes(AWriter);
    if xaElements in FSpPr.Assigneds then
    begin
      AWriter.BeginTag('c:spPr');
      FSpPr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:spPr');
  end;
{$ifdef XLS_WRITE_EXTLST}
  if (FExtLst <> Nil) and FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then
    begin
      AWriter.BeginTag('c:extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:extLst');
{$endif}
end;

constructor TCT_PlotArea.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 24;
  FAttributeCount := 0;
end;

destructor TCT_PlotArea.Destroy;
begin
  if FLayout <> Nil then 
    FLayout.Free;
  if FAreaChartXpgList <> Nil then 
    FAreaChartXpgList.Free;
  if FArea3DChartXpgList <> Nil then 
    FArea3DChartXpgList.Free;
  if FLineChartXpgList <> Nil then 
    FLineChartXpgList.Free;
  if FLine3DChartXpgList <> Nil then 
    FLine3DChartXpgList.Free;
  if FStockChartXpgList <> Nil then 
    FStockChartXpgList.Free;
  if FRadarChartXpgList <> Nil then 
    FRadarChartXpgList.Free;
  if FScatterChartXpgList <> Nil then 
    FScatterChartXpgList.Free;
  if FPieChartXpgList <> Nil then 
    FPieChartXpgList.Free;
  if FPie3DChartXpgList <> Nil then 
    FPie3DChartXpgList.Free;
  if FDoughnutChartXpgList <> Nil then 
    FDoughnutChartXpgList.Free;
  if FBarChartXpgList <> Nil then 
    FBarChartXpgList.Free;
  if FBar3DChartXpgList <> Nil then 
    FBar3DChartXpgList.Free;
  if FOfPieChartXpgList <> Nil then 
    FOfPieChartXpgList.Free;
  if FSurfaceChartXpgList <> Nil then 
    FSurfaceChartXpgList.Free;
  if FSurface3DChartXpgList <> Nil then 
    FSurface3DChartXpgList.Free;
  if FBubbleChartXpgList <> Nil then 
    FBubbleChartXpgList.Free;
  if FValAxXpgList <> Nil then 
    FValAxXpgList.Free;
  if FCatAxXpgList <> Nil then 
    FCatAxXpgList.Free;
  if FDateAxXpgList <> Nil then 
    FDateAxXpgList.Free;
  if FSerAxXpgList <> Nil then 
    FSerAxXpgList.Free;
  if FDTable <> Nil then 
    FDTable.Free;
  if FSpPr <> Nil then 
    FSpPr.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_PlotArea.Clear;
begin
  FAssigneds := [];
  if FLayout <> Nil then 
    FreeAndNil(FLayout);
  if FAreaChartXpgList <> Nil then 
    FreeAndNil(FAreaChartXpgList);
  if FArea3DChartXpgList <> Nil then 
    FreeAndNil(FArea3DChartXpgList);
  if FLineChartXpgList <> Nil then 
    FreeAndNil(FLineChartXpgList);
  if FLine3DChartXpgList <> Nil then 
    FreeAndNil(FLine3DChartXpgList);
  if FStockChartXpgList <> Nil then 
    FreeAndNil(FStockChartXpgList);
  if FRadarChartXpgList <> Nil then 
    FreeAndNil(FRadarChartXpgList);
  if FScatterChartXpgList <> Nil then 
    FreeAndNil(FScatterChartXpgList);
  if FPieChartXpgList <> Nil then 
    FreeAndNil(FPieChartXpgList);
  if FPie3DChartXpgList <> Nil then 
    FreeAndNil(FPie3DChartXpgList);
  if FDoughnutChartXpgList <> Nil then 
    FreeAndNil(FDoughnutChartXpgList);
  if FBarChartXpgList <> Nil then 
    FreeAndNil(FBarChartXpgList);
  if FBar3DChartXpgList <> Nil then 
    FreeAndNil(FBar3DChartXpgList);
  if FOfPieChartXpgList <> Nil then 
    FreeAndNil(FOfPieChartXpgList);
  if FSurfaceChartXpgList <> Nil then 
    FreeAndNil(FSurfaceChartXpgList);
  if FSurface3DChartXpgList <> Nil then 
    FreeAndNil(FSurface3DChartXpgList);
  if FBubbleChartXpgList <> Nil then 
    FreeAndNil(FBubbleChartXpgList);
  if FValAxXpgList <> Nil then 
    FreeAndNil(FValAxXpgList);
  if FCatAxXpgList <> Nil then 
    FreeAndNil(FCatAxXpgList);
  if FDateAxXpgList <> Nil then 
    FreeAndNil(FDateAxXpgList);
  if FSerAxXpgList <> Nil then 
    FreeAndNil(FSerAxXpgList);
  if FDTable <> Nil then 
    FreeAndNil(FDTable);
  if FSpPr <> Nil then 
    FreeAndNil(FSpPr);
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
end;

procedure TCT_PlotArea.Assign(AItem: TCT_PlotArea);
begin
end;

procedure TCT_PlotArea.CopyTo(AItem: TCT_PlotArea);
begin
end;

function  TCT_PlotArea.Create_Layout: TCT_Layout;
begin
  Result := TCT_Layout.Create(FOwner);
  FLayout := Result;
end;

function  TCT_PlotArea.Create_AreaChartXpgList: TCT_AreaChartXpgList;
begin
  Result := TCT_AreaChartXpgList.Create(FOwner);
  FAreaChartXpgList := Result;
end;

function  TCT_PlotArea.Create_Area3DChartXpgList: TCT_Area3DChartXpgList;
begin
  Result := TCT_Area3DChartXpgList.Create(FOwner);
  FArea3DChartXpgList := Result;
end;

function  TCT_PlotArea.Create_LineChartXpgList: TCT_LineChartXpgList;
begin
  Result := TCT_LineChartXpgList.Create(FOwner);
  FLineChartXpgList := Result;
end;

function  TCT_PlotArea.Create_Line3DChartXpgList: TCT_Line3DChartXpgList;
begin
  Result := TCT_Line3DChartXpgList.Create(FOwner);
  FLine3DChartXpgList := Result;
end;

function  TCT_PlotArea.Create_StockChartXpgList: TCT_StockChartXpgList;
begin
  Result := TCT_StockChartXpgList.Create(FOwner);
  FStockChartXpgList := Result;
end;

function  TCT_PlotArea.Create_RadarChartXpgList: TCT_RadarChartXpgList;
begin
  Result := TCT_RadarChartXpgList.Create(FOwner);
  FRadarChartXpgList := Result;
end;

function  TCT_PlotArea.Create_ScatterChartXpgList: TCT_ScatterChartXpgList;
begin
  Result := TCT_ScatterChartXpgList.Create(FOwner);
  FScatterChartXpgList := Result;
end;

function  TCT_PlotArea.Create_PieChartXpgList: TCT_PieChartXpgList;
begin
  Result := TCT_PieChartXpgList.Create(FOwner);
  FPieChartXpgList := Result;
end;

function  TCT_PlotArea.Create_Pie3DChartXpgList: TCT_Pie3DChartXpgList;
begin
  Result := TCT_Pie3DChartXpgList.Create(FOwner);
  FPie3DChartXpgList := Result;
end;

function  TCT_PlotArea.Create_DoughnutChartXpgList: TCT_DoughnutChartXpgList;
begin
  Result := TCT_DoughnutChartXpgList.Create(FOwner);
  FDoughnutChartXpgList := Result;
end;

function  TCT_PlotArea.Create_BarChartXpgList: TCT_BarChartXpgList;
begin
  Result := TCT_BarChartXpgList.Create(FOwner);
  FBarChartXpgList := Result;
end;

function  TCT_PlotArea.Create_Bar3DChartXpgList: TCT_Bar3DChartXpgList;
begin
  Result := TCT_Bar3DChartXpgList.Create(FOwner);
  FBar3DChartXpgList := Result;
end;

function  TCT_PlotArea.Create_OfPieChartXpgList: TCT_OfPieChartXpgList;
begin
  Result := TCT_OfPieChartXpgList.Create(FOwner);
  FOfPieChartXpgList := Result;
end;

function  TCT_PlotArea.Create_SurfaceChartXpgList: TCT_SurfaceChartXpgList;
begin
  Result := TCT_SurfaceChartXpgList.Create(FOwner);
  FSurfaceChartXpgList := Result;
end;

function  TCT_PlotArea.Create_Surface3DChartXpgList: TCT_Surface3DChartXpgList;
begin
  Result := TCT_Surface3DChartXpgList.Create(FOwner);
  FSurface3DChartXpgList := Result;
end;

function  TCT_PlotArea.Create_BubbleChartXpgList: TCT_BubbleChartXpgList;
begin
  Result := TCT_BubbleChartXpgList.Create(FOwner);
  FBubbleChartXpgList := Result;
end;

function  TCT_PlotArea.Create_ValAxXpgList: TCT_ValAxXpgList;
begin
  Result := TCT_ValAxXpgList.Create(FOwner);
  FValAxXpgList := Result;
end;

function  TCT_PlotArea.Create_CatAxXpgList: TCT_CatAxXpgList;
begin
  Result := TCT_CatAxXpgList.Create(FOwner);
  FCatAxXpgList := Result;
end;

function  TCT_PlotArea.Create_DateAxXpgList: TCT_DateAxXpgList;
begin
  Result := TCT_DateAxXpgList.Create(FOwner);
  FDateAxXpgList := Result;
end;

function  TCT_PlotArea.Create_SerAxXpgList: TCT_SerAxXpgList;
begin
  Result := TCT_SerAxXpgList.Create(FOwner);
  FSerAxXpgList := Result;
end;

function  TCT_PlotArea.Create_DTable: TCT_DTable;
begin
  Result := TCT_DTable.Create(FOwner);
  FDTable := Result;
end;

function  TCT_PlotArea.Create_SpPr: TCT_ShapeProperties;
begin
  Result := TCT_ShapeProperties.Create(FOwner);
  FSpPr := Result;
end;

function  TCT_PlotArea.Create_ExtLst: TCT_ExtensionList;
begin
  Result := TCT_ExtensionList.Create(FOwner);
  FExtLst := Result;
end;

{ TCT_Legend }

function  TCT_Legend.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FLegendPos <> Nil then 
    Inc(ElemsAssigned,FLegendPos.CheckAssigned);
  if FLegendEntryXpgList <> Nil then 
    Inc(ElemsAssigned,FLegendEntryXpgList.CheckAssigned);
  if FLayout <> Nil then 
    Inc(ElemsAssigned,FLayout.CheckAssigned);
  if FOverlay <> Nil then 
    Inc(ElemsAssigned,FOverlay.CheckAssigned);
  if FSpPr <> Nil then 
    Inc(ElemsAssigned,FSpPr.CheckAssigned);
  if FTxPr <> Nil then 
    Inc(ElemsAssigned,FTxPr.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_Legend.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $0000043E: begin
      if FLegendPos = Nil then 
        FLegendPos := TCT_LegendPos.Create(FOwner);
      Result := FLegendPos;
    end;
    $0000051E: begin
      if FLegendEntryXpgList = Nil then 
        FLegendEntryXpgList := TCT_LegendEntryXpgList.Create(FOwner);
      Result := FLegendEntryXpgList.Add;
    end;
    $0000033B: begin
      if FLayout = Nil then 
        FLayout := TCT_Layout.Create(FOwner);
      Result := FLayout;
    end;
    $0000039F: begin
      if FOverlay = Nil then 
        FOverlay := TCT_Boolean.Create(FOwner);
      Result := FOverlay;
    end;
    $00000242: begin
      if FSpPr = Nil then 
        FSpPr := TCT_ShapeProperties.Create(FOwner);
      Result := FSpPr;
    end;
    $0000024B: begin
      if FTxPr = Nil then 
        FTxPr := TCT_TextBody.Create(FOwner);
      Result := FTxPr;
    end;
    $00000321: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_Legend.Write(AWriter: TXpgWriteXML);
begin
//  if (FLegendPos <> Nil) and FLegendPos.Assigned then
//  begin
    FLegendPos.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:legendPos');
//  end;
  if FLegendEntryXpgList <> Nil then 
    FLegendEntryXpgList.Write(AWriter,'c:legendEntry');
  if (FLayout <> Nil) and FLayout.Assigned then 
    if xaElements in FLayout.Assigneds then
    begin
      AWriter.BeginTag('c:layout');
      FLayout.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:layout');
  if (FOverlay <> Nil) and FOverlay.Assigned then 
  begin
    FOverlay.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:overlay');
  end;
  if (FSpPr <> Nil) and FSpPr.Assigned then 
  begin
    FSpPr.WriteAttributes(AWriter);
    if xaElements in FSpPr.Assigneds then
    begin
      AWriter.BeginTag('c:spPr');
      FSpPr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:spPr');
  end;
  if (FTxPr <> Nil) and FTxPr.Assigned then 
    if xaElements in FTxPr.Assigneds then
    begin
      AWriter.BeginTag('c:txPr');
      FTxPr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:txPr');
{$ifdef XLS_WRITE_EXTLST}
  if (FExtLst <> Nil) and FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then
    begin
      AWriter.BeginTag('c:extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:extLst');
{$endif}
end;

constructor TCT_Legend.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 7;
  FAttributeCount := 0;
end;

destructor TCT_Legend.Destroy;
begin
  if FLegendPos <> Nil then 
    FLegendPos.Free;
  if FLegendEntryXpgList <> Nil then 
    FLegendEntryXpgList.Free;
  if FLayout <> Nil then 
    FLayout.Free;
  if FOverlay <> Nil then 
    FOverlay.Free;
  if FSpPr <> Nil then 
    FSpPr.Free;
  if FTxPr <> Nil then 
    FTxPr.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_Legend.Clear;
begin
  FAssigneds := [];
  if FLegendPos <> Nil then 
    FreeAndNil(FLegendPos);
  if FLegendEntryXpgList <> Nil then 
    FreeAndNil(FLegendEntryXpgList);
  if FLayout <> Nil then 
    FreeAndNil(FLayout);
  if FOverlay <> Nil then 
    FreeAndNil(FOverlay);
  if FSpPr <> Nil then 
    FreeAndNil(FSpPr);
  if FTxPr <> Nil then 
    FreeAndNil(FTxPr);
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
end;

procedure TCT_Legend.Assign(AItem: TCT_Legend);
begin
end;

procedure TCT_Legend.CopyTo(AItem: TCT_Legend);
begin
end;

function  TCT_Legend.Create_LegendPos: TCT_LegendPos;
begin
  Result := TCT_LegendPos.Create(FOwner);
  FLegendPos := Result;
end;

function  TCT_Legend.Create_LegendEntryXpgList: TCT_LegendEntryXpgList;
begin
  Result := TCT_LegendEntryXpgList.Create(FOwner);
  FLegendEntryXpgList := Result;
end;

function  TCT_Legend.Create_Layout: TCT_Layout;
begin
  Result := TCT_Layout.Create(FOwner);
  FLayout := Result;
end;

function  TCT_Legend.Create_Overlay: TCT_Boolean;
begin
  Result := TCT_Boolean.Create(FOwner);
  FOverlay := Result;
end;

function  TCT_Legend.Create_SpPr: TCT_ShapeProperties;
begin
  Result := TCT_ShapeProperties.Create(FOwner);
  FSpPr := Result;
end;

function  TCT_Legend.Create_TxPr: TCT_TextBody;
begin
  Result := TCT_TextBody.Create(FOwner);
  FTxPr := Result;
end;

function  TCT_Legend.Create_ExtLst: TCT_ExtensionList;
begin
  Result := TCT_ExtensionList.Create(FOwner);
  FExtLst := Result;
end;

{ TCT_DispBlanksAs }

function  TCT_DispBlanksAs.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FVal <> stdbaZero then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_DispBlanksAs.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_DispBlanksAs.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FVal <> stdbaZero then 
    AWriter.AddAttribute('val',StrTST_DispBlanksAs[Integer(FVal)]);
end;

procedure TCT_DispBlanksAs.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then 
    FVal := TST_DispBlanksAs(StrToEnum('stdba' + AAttributes.Values[0]))
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_DispBlanksAs.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := stdbaZero;
end;

destructor TCT_DispBlanksAs.Destroy;
begin
end;

procedure TCT_DispBlanksAs.Clear;
begin
  FAssigneds := [];
  FVal := stdbaZero;
end;

procedure TCT_DispBlanksAs.Assign(AItem: TCT_DispBlanksAs);
begin
end;

procedure TCT_DispBlanksAs.CopyTo(AItem: TCT_DispBlanksAs);
begin
end;

{ TCT_HeaderFooter }

function  TCT_HeaderFooter.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FAlignWithMargins <> True then 
    Inc(AttrsAssigned);
  if FDifferentOddEven <> False then 
    Inc(AttrsAssigned);
  if FDifferentFirst <> False then 
    Inc(AttrsAssigned);
  if FOddHeader <> '' then 
    Inc(ElemsAssigned);
  if FOddFooter <> '' then 
    Inc(ElemsAssigned);
  if FEvenHeader <> '' then 
    Inc(ElemsAssigned);
  if FEvenFooter <> '' then 
    Inc(ElemsAssigned);
  if FFirstHeader <> '' then 
    Inc(ElemsAssigned);
  if FFirstFooter <> '' then 
    Inc(ElemsAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_HeaderFooter.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $0000041D: FOddHeader := AReader.Text;
    $00000443: FOddFooter := AReader.Text;
    $00000494: FEvenHeader := AReader.Text;
    $000004BA: FEvenFooter := AReader.Text;
    $0000050E: FFirstHeader := AReader.Text;
    $00000534: FFirstFooter := AReader.Text;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_HeaderFooter.Write(AWriter: TXpgWriteXML);
begin
  if FOddHeader <> '' then 
    AWriter.SimpleTextTag('c:oddHeader',FOddHeader);
  if FOddFooter <> '' then 
    AWriter.SimpleTextTag('c:oddFooter',FOddFooter);
  if FEvenHeader <> '' then 
    AWriter.SimpleTextTag('c:evenHeader',FEvenHeader);
  if FEvenFooter <> '' then 
    AWriter.SimpleTextTag('c:evenFooter',FEvenFooter);
  if FFirstHeader <> '' then 
    AWriter.SimpleTextTag('c:firstHeader',FFirstHeader);
  if FFirstFooter <> '' then 
    AWriter.SimpleTextTag('c:firstFooter',FFirstFooter);
end;

procedure TCT_HeaderFooter.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FAlignWithMargins <> True then 
    AWriter.AddAttribute('alignWithMargins',XmlBoolToStr(FAlignWithMargins));
  if FDifferentOddEven <> False then 
    AWriter.AddAttribute('differentOddEven',XmlBoolToStr(FDifferentOddEven));
  if FDifferentFirst <> False then 
    AWriter.AddAttribute('differentFirst',XmlBoolToStr(FDifferentFirst));
end;

procedure TCT_HeaderFooter.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $00000678: FAlignWithMargins := XmlStrToBoolDef(AAttributes.Values[i],True);
      $0000065C: FDifferentOddEven := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000005BF: FDifferentFirst := XmlStrToBoolDef(AAttributes.Values[i],False);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_HeaderFooter.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 6;
  FAttributeCount := 3;
  FAlignWithMargins := True;
  FDifferentOddEven := False;
  FDifferentFirst := False;
end;

destructor TCT_HeaderFooter.Destroy;
begin
end;

procedure TCT_HeaderFooter.Clear;
begin
  FAssigneds := [];
  FOddHeader := '';
  FOddFooter := '';
  FEvenHeader := '';
  FEvenFooter := '';
  FFirstHeader := '';
  FFirstFooter := '';
  FAlignWithMargins := True;
  FDifferentOddEven := False;
  FDifferentFirst := False;
end;

procedure TCT_HeaderFooter.Assign(AItem: TCT_HeaderFooter);
begin
end;

procedure TCT_HeaderFooter.CopyTo(AItem: TCT_HeaderFooter);
begin
end;

{ TCT_PageMargins }

function  TCT_PageMargins.CheckAssigned: integer;
begin
  FAssigneds := [xaAttributes];
  Result := 1;
end;

procedure TCT_PageMargins.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_PageMargins.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('l',XmlFloatToStr(FL));
  AWriter.AddAttribute('r',XmlFloatToStr(FR));
  AWriter.AddAttribute('t',XmlFloatToStr(FT));
  AWriter.AddAttribute('b',XmlFloatToStr(FB));
  AWriter.AddAttribute('header',XmlFloatToStr(FHeader));
  AWriter.AddAttribute('footer',XmlFloatToStr(FFooter));
end;

procedure TCT_PageMargins.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $0000006C: FL := XmlStrToFloatDef(AAttributes.Values[i],0);
      $00000072: FR := XmlStrToFloatDef(AAttributes.Values[i],0);
      $00000074: FT := XmlStrToFloatDef(AAttributes.Values[i],0);
      $00000062: FB := XmlStrToFloatDef(AAttributes.Values[i],0);
      $00000269: FHeader := XmlStrToFloatDef(AAttributes.Values[i],0);
      $0000028F: FFooter := XmlStrToFloatDef(AAttributes.Values[i],0);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_PageMargins.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 6;
  FL := NaN;
  FR := NaN;
  FT := NaN;
  FB := NaN;
  FHeader := NaN;
  FFooter := NaN;
end;

destructor TCT_PageMargins.Destroy;
begin
end;

procedure TCT_PageMargins.Clear;
begin
  FAssigneds := [];
  FL := NaN;
  FR := NaN;
  FT := NaN;
  FB := NaN;
  FHeader := NaN;
  FFooter := NaN;
end;

procedure TCT_PageMargins.Assign(AItem: TCT_PageMargins);
begin
end;

procedure TCT_PageMargins.CopyTo(AItem: TCT_PageMargins);
begin
end;

{ TCT_PageSetup }

function  TCT_PageSetup.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FPaperSize <> 1 then 
    Inc(AttrsAssigned);
  if FFirstPageNumber <> 1 then 
    Inc(AttrsAssigned);
  if FOrientation <> stpsoDefault then 
    Inc(AttrsAssigned);
  if FBlackAndWhite <> False then 
    Inc(AttrsAssigned);
  if FDraft <> False then 
    Inc(AttrsAssigned);
  if FUseFirstPageNumber <> False then 
    Inc(AttrsAssigned);
  if FHorizontalDpi <> 600 then 
    Inc(AttrsAssigned);
  if FVerticalDpi <> 600 then 
    Inc(AttrsAssigned);
  if FCopies <> 1 then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_PageSetup.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_PageSetup.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FPaperSize <> 1 then 
    AWriter.AddAttribute('paperSize',XmlIntToStr(FPaperSize));
  if FFirstPageNumber <> 1 then 
    AWriter.AddAttribute('firstPageNumber',XmlIntToStr(FFirstPageNumber));
  if FOrientation <> stpsoDefault then 
    AWriter.AddAttribute('orientation',StrTST_PageSetupOrientation[Integer(FOrientation)]);
  if FBlackAndWhite <> False then 
    AWriter.AddAttribute('blackAndWhite',XmlBoolToStr(FBlackAndWhite));
  if FDraft <> False then 
    AWriter.AddAttribute('draft',XmlBoolToStr(FDraft));
  if FUseFirstPageNumber <> False then 
    AWriter.AddAttribute('useFirstPageNumber',XmlBoolToStr(FUseFirstPageNumber));
  if FHorizontalDpi <> 600 then 
    AWriter.AddAttribute('horizontalDpi',XmlIntToStr(FHorizontalDpi));
  if FVerticalDpi <> 600 then 
    AWriter.AddAttribute('verticalDpi',XmlIntToStr(FVerticalDpi));
  if FCopies <> 1 then 
    AWriter.AddAttribute('copies',XmlIntToStr(FCopies));
end;

procedure TCT_PageSetup.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $000003B3: FPaperSize := XmlStrToIntDef(AAttributes.Values[i],0);
      $0000060E: FFirstPageNumber := XmlStrToIntDef(AAttributes.Values[i],0);
      $000004AC: FOrientation := TST_PageSetupOrientation(StrToEnum('stpso' + AAttributes.Values[i]));
      $00000511: FBlackAndWhite := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000211: FDraft := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000073B: FUseFirstPageNumber := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000567: FHorizontalDpi := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000477: FVerticalDpi := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000283: FCopies := XmlStrToIntDef(AAttributes.Values[i],0);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_PageSetup.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 9;
  Clear;
end;

destructor TCT_PageSetup.Destroy;
begin
end;

procedure TCT_PageSetup.Clear;
begin
  FAssigneds := [];
  FPaperSize := 1;
  FFirstPageNumber := 1;
  FOrientation := stpsoDefault;
  FBlackAndWhite := False;
  FDraft := False;
  FUseFirstPageNumber := False;
  FHorizontalDpi := 600;
  FVerticalDpi := 600;
  FCopies := 1;
end;

procedure TCT_PageSetup.Assign(AItem: TCT_PageSetup);
begin
end;

procedure TCT_PageSetup.CopyTo(AItem: TCT_PageSetup);
begin
end;

{ TCT_RelId }

function  TCT_RelId.CheckAssigned: integer;
begin
  FAssigneds := [xaAttributes];
  Result := 1;
end;

procedure TCT_RelId.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_RelId.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('r:id',FR_Id);
end;

procedure TCT_RelId.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'r:id' then 
    FR_Id := AAttributes.Values[0]
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_RelId.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
end;

destructor TCT_RelId.Destroy;
begin
end;

procedure TCT_RelId.Clear;
begin
  FAssigneds := [];
  FR_Id := '';
end;

procedure TCT_RelId.Assign(AItem: TCT_RelId);
begin
end;

procedure TCT_RelId.CopyTo(AItem: TCT_RelId);
begin
end;

{ TCT_ObjectStyleDefaults }

function  TCT_ObjectStyleDefaults.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FA_SpDef <> Nil then 
    Inc(ElemsAssigned,FA_SpDef.CheckAssigned);
  if FA_LnDef <> Nil then 
    Inc(ElemsAssigned,FA_LnDef.CheckAssigned);
  if FA_TxDef <> Nil then 
    Inc(ElemsAssigned,FA_TxDef.CheckAssigned);
  if FA_ExtLst <> Nil then 
    Inc(ElemsAssigned,FA_ExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_ObjectStyleDefaults.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $0000028D: begin
      if FA_SpDef = Nil then 
        FA_SpDef := TCT_DefaultShapeDefinition.Create(FOwner);
      Result := FA_SpDef;
    end;
    $00000284: begin
      if FA_LnDef = Nil then 
        FA_LnDef := TCT_DefaultShapeDefinition.Create(FOwner);
      Result := FA_LnDef;
    end;
    $00000296: begin
      if FA_TxDef = Nil then 
        FA_TxDef := TCT_DefaultShapeDefinition.Create(FOwner);
      Result := FA_TxDef;
    end;
    $0000031F: begin
      if FA_ExtLst = Nil then 
        FA_ExtLst := TCT_OfficeArtExtensionList.Create(FOwner);
      Result := FA_ExtLst;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_ObjectStyleDefaults.Write(AWriter: TXpgWriteXML);
begin
  if (FA_SpDef <> Nil) and FA_SpDef.Assigned then 
    if xaElements in FA_SpDef.Assigneds then
    begin
      AWriter.BeginTag('a:spDef');
      FA_SpDef.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:spDef');
  if (FA_LnDef <> Nil) and FA_LnDef.Assigned then 
    if xaElements in FA_LnDef.Assigneds then
    begin
      AWriter.BeginTag('a:lnDef');
      FA_LnDef.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:lnDef');
  if (FA_TxDef <> Nil) and FA_TxDef.Assigned then 
    if xaElements in FA_TxDef.Assigneds then
    begin
      AWriter.BeginTag('a:txDef');
      FA_TxDef.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:txDef');
  if (FA_ExtLst <> Nil) and FA_ExtLst.Assigned then 
    if xaElements in FA_ExtLst.Assigneds then
    begin
      AWriter.BeginTag('a:extLst');
      FA_ExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:extLst');
end;

constructor TCT_ObjectStyleDefaults.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 4;
  FAttributeCount := 0;
end;

destructor TCT_ObjectStyleDefaults.Destroy;
begin
  if FA_SpDef <> Nil then 
    FA_SpDef.Free;
  if FA_LnDef <> Nil then 
    FA_LnDef.Free;
  if FA_TxDef <> Nil then 
    FA_TxDef.Free;
  if FA_ExtLst <> Nil then 
    FA_ExtLst.Free;
end;

procedure TCT_ObjectStyleDefaults.Clear;
begin
  FAssigneds := [];
  if FA_SpDef <> Nil then 
    FreeAndNil(FA_SpDef);
  if FA_LnDef <> Nil then 
    FreeAndNil(FA_LnDef);
  if FA_TxDef <> Nil then 
    FreeAndNil(FA_TxDef);
  if FA_ExtLst <> Nil then 
    FreeAndNil(FA_ExtLst);
end;

procedure TCT_ObjectStyleDefaults.Assign(AItem: TCT_ObjectStyleDefaults);
begin
end;

procedure TCT_ObjectStyleDefaults.CopyTo(AItem: TCT_ObjectStyleDefaults);
begin
end;

function  TCT_ObjectStyleDefaults.Create_A_SpDef: TCT_DefaultShapeDefinition;
begin
  Result := TCT_DefaultShapeDefinition.Create(FOwner);
  FA_SpDef := Result;
end;

function  TCT_ObjectStyleDefaults.Create_A_LnDef: TCT_DefaultShapeDefinition;
begin
  Result := TCT_DefaultShapeDefinition.Create(FOwner);
  FA_LnDef := Result;
end;

function  TCT_ObjectStyleDefaults.Create_A_TxDef: TCT_DefaultShapeDefinition;
begin
  Result := TCT_DefaultShapeDefinition.Create(FOwner);
  FA_TxDef := Result;
end;

function  TCT_ObjectStyleDefaults.Create_A_ExtLst: TCT_OfficeArtExtensionList;
begin
  Result := TCT_OfficeArtExtensionList.Create(FOwner);
  FA_ExtLst := Result;
end;

{ TCT_ColorSchemeList }

function  TCT_ColorSchemeList.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FA_ExtraClrSchemeXpgList <> Nil then 
    Inc(ElemsAssigned,FA_ExtraClrSchemeXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_ColorSchemeList.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'a:extraClrScheme' then 
  begin
    if FA_ExtraClrSchemeXpgList = Nil then 
      FA_ExtraClrSchemeXpgList := TCT_ColorSchemeAndMappingXpgList.Create(FOwner);
    Result := FA_ExtraClrSchemeXpgList.Add;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_ColorSchemeList.Write(AWriter: TXpgWriteXML);
begin
  if FA_ExtraClrSchemeXpgList <> Nil then 
    FA_ExtraClrSchemeXpgList.Write(AWriter,'a:extraClrScheme');
end;

constructor TCT_ColorSchemeList.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
end;

destructor TCT_ColorSchemeList.Destroy;
begin
  if FA_ExtraClrSchemeXpgList <> Nil then 
    FA_ExtraClrSchemeXpgList.Free;
end;

procedure TCT_ColorSchemeList.Clear;
begin
  FAssigneds := [];
  if FA_ExtraClrSchemeXpgList <> Nil then 
    FreeAndNil(FA_ExtraClrSchemeXpgList);
end;

procedure TCT_ColorSchemeList.Assign(AItem: TCT_ColorSchemeList);
begin
end;

procedure TCT_ColorSchemeList.CopyTo(AItem: TCT_ColorSchemeList);
begin
end;

function  TCT_ColorSchemeList.Create_A_ExtraClrSchemeXpgList: TCT_ColorSchemeAndMappingXpgList;
begin
  Result := TCT_ColorSchemeAndMappingXpgList.Create(FOwner);
  FA_ExtraClrSchemeXpgList := Result;
end;

{ TCT_TextLanguageID }

function  TCT_TextLanguageID.CheckAssigned: integer;
begin
  FAssigneds := [xaAttributes];
  Result := 1;
end;

procedure TCT_TextLanguageID.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_TextLanguageID.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('val',FVal);
end;

procedure TCT_TextLanguageID.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then 
    FVal := AAttributes.Values[0]
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_TextLanguageID.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
end;

destructor TCT_TextLanguageID.Destroy;
begin
end;

procedure TCT_TextLanguageID.Clear;
begin
  FAssigneds := [];
  FVal := '';
end;

procedure TCT_TextLanguageID.Assign(AItem: TCT_TextLanguageID);
begin
end;

procedure TCT_TextLanguageID.CopyTo(AItem: TCT_TextLanguageID);
begin
end;

{ TCT_Style }

function  TCT_Style.CheckAssigned: integer;
begin
  FAssigneds := [xaAttributes];
  Result := 1;
end;

procedure TCT_Style.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_Style.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('val',XmlIntToStr(FVal));
end;

procedure TCT_Style.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then 
    FVal := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_Style.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := 2147483632;
end;

destructor TCT_Style.Destroy;
begin
end;

procedure TCT_Style.Clear;
begin
  FAssigneds := [];
  FVal := 2147483632;
end;

procedure TCT_Style.Assign(AItem: TCT_Style);
begin
end;

procedure TCT_Style.CopyTo(AItem: TCT_Style);
begin
end;

{ TCT_PivotSource }

function  TCT_PivotSource.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FName <> '' then 
    Inc(ElemsAssigned);
  if FFmtId <> Nil then 
    Inc(ElemsAssigned,FFmtId.CheckAssigned);
  if FExtLstXpgList <> Nil then 
    Inc(ElemsAssigned,FExtLstXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_PivotSource.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $0000023E: FName := AReader.Text;
    $00000291: begin
      if FFmtId = Nil then 
        FFmtId := TCT_UnsignedInt.Create(FOwner);
      Result := FFmtId;
    end;
    $00000321: begin
      if FExtLstXpgList = Nil then 
        FExtLstXpgList := TCT_ExtensionListXpgList.Create(FOwner);
      Result := FExtLstXpgList.Add;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_PivotSource.Write(AWriter: TXpgWriteXML);
begin
  AWriter.SimpleTextTag('c:name',FName);
  if (FFmtId <> Nil) and FFmtId.Assigned then 
  begin
    FFmtId.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:fmtId');
  end
  else 
    AWriter.SimpleTag('c:fmtId');
  if FExtLstXpgList <> Nil then 
    FExtLstXpgList.Write(AWriter,'c:extLst');
end;

constructor TCT_PivotSource.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 3;
  FAttributeCount := 0;
end;

destructor TCT_PivotSource.Destroy;
begin
  if FFmtId <> Nil then 
    FFmtId.Free;
  if FExtLstXpgList <> Nil then 
    FExtLstXpgList.Free;
end;

procedure TCT_PivotSource.Clear;
begin
  FAssigneds := [];
  FName := '';
  if FFmtId <> Nil then 
    FreeAndNil(FFmtId);
  if FExtLstXpgList <> Nil then 
    FreeAndNil(FExtLstXpgList);
end;

procedure TCT_PivotSource.Assign(AItem: TCT_PivotSource);
begin
end;

procedure TCT_PivotSource.CopyTo(AItem: TCT_PivotSource);
begin
end;

function  TCT_PivotSource.Create_FmtId: TCT_UnsignedInt;
begin
  Result := TCT_UnsignedInt.Create(FOwner);
  FFmtId := Result;
end;

function  TCT_PivotSource.Create_ExtLstXpgList: TCT_ExtensionListXpgList;
begin
  Result := TCT_ExtensionListXpgList.Create(FOwner);
  FExtLstXpgList := Result;
end;

{ TCT_Protection }

function  TCT_Protection.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FChartObject <> Nil then 
    Inc(ElemsAssigned,FChartObject.CheckAssigned);
  if FData <> Nil then 
    Inc(ElemsAssigned,FData.CheckAssigned);
  if FFormatting <> Nil then 
    Inc(ElemsAssigned,FFormatting.CheckAssigned);
  if FSelection <> Nil then 
    Inc(ElemsAssigned,FSelection.CheckAssigned);
  if FUserInterface <> Nil then 
    Inc(ElemsAssigned,FUserInterface.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_Protection.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $00000506: begin
      if FChartObject = Nil then 
        FChartObject := TCT_Boolean.Create(FOwner);
      Result := FChartObject;
    end;
    $00000237: begin
      if FData = Nil then 
        FData := TCT_Boolean.Create(FOwner);
      Result := FData;
    end;
    $000004D8: begin
      if FFormatting = Nil then 
        FFormatting := TCT_Boolean.Create(FOwner);
      Result := FFormatting;
    end;
    $00000463: begin
      if FSelection = Nil then 
        FSelection := TCT_Boolean.Create(FOwner);
      Result := FSelection;
    end;
    $000005ED: begin
      if FUserInterface = Nil then 
        FUserInterface := TCT_Boolean.Create(FOwner);
      Result := FUserInterface;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_Protection.Write(AWriter: TXpgWriteXML);
begin
  if (FChartObject <> Nil) and FChartObject.Assigned then 
  begin
    FChartObject.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:chartObject');
  end;
  if (FData <> Nil) and FData.Assigned then 
  begin
    FData.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:data');
  end;
  if (FFormatting <> Nil) and FFormatting.Assigned then 
  begin
    FFormatting.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:formatting');
  end;
  if (FSelection <> Nil) and FSelection.Assigned then 
  begin
    FSelection.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:selection');
  end;
  if (FUserInterface <> Nil) and FUserInterface.Assigned then 
  begin
    FUserInterface.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:userInterface');
  end;
end;

constructor TCT_Protection.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 5;
  FAttributeCount := 0;
end;

destructor TCT_Protection.Destroy;
begin
  if FChartObject <> Nil then 
    FChartObject.Free;
  if FData <> Nil then 
    FData.Free;
  if FFormatting <> Nil then 
    FFormatting.Free;
  if FSelection <> Nil then 
    FSelection.Free;
  if FUserInterface <> Nil then 
    FUserInterface.Free;
end;

procedure TCT_Protection.Clear;
begin
  FAssigneds := [];
  if FChartObject <> Nil then 
    FreeAndNil(FChartObject);
  if FData <> Nil then 
    FreeAndNil(FData);
  if FFormatting <> Nil then 
    FreeAndNil(FFormatting);
  if FSelection <> Nil then 
    FreeAndNil(FSelection);
  if FUserInterface <> Nil then 
    FreeAndNil(FUserInterface);
end;

procedure TCT_Protection.Assign(AItem: TCT_Protection);
begin
end;

procedure TCT_Protection.CopyTo(AItem: TCT_Protection);
begin
end;

function  TCT_Protection.Create_ChartObject: TCT_Boolean;
begin
  Result := TCT_Boolean.Create(FOwner);
  FChartObject := Result;
end;

function  TCT_Protection.Create_Data: TCT_Boolean;
begin
  Result := TCT_Boolean.Create(FOwner);
  FData := Result;
end;

function  TCT_Protection.Create_Formatting: TCT_Boolean;
begin
  Result := TCT_Boolean.Create(FOwner);
  FFormatting := Result;
end;

function  TCT_Protection.Create_Selection: TCT_Boolean;
begin
  Result := TCT_Boolean.Create(FOwner);
  FSelection := Result;
end;

function  TCT_Protection.Create_UserInterface: TCT_Boolean;
begin
  Result := TCT_Boolean.Create(FOwner);
  FUserInterface := Result;
end;

{ TCT_Chart }

function  TCT_Chart.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FTitle <> Nil then 
    Inc(ElemsAssigned,FTitle.CheckAssigned);
  if FAutoTitleDeleted <> Nil then 
    Inc(ElemsAssigned,FAutoTitleDeleted.CheckAssigned);
  if FPivotFmts <> Nil then 
    Inc(ElemsAssigned,FPivotFmts.CheckAssigned);
  if FView3D <> Nil then 
    Inc(ElemsAssigned,FView3D.CheckAssigned);
  if FFloor <> Nil then 
    Inc(ElemsAssigned,FFloor.CheckAssigned);
  if FSideWall <> Nil then 
    Inc(ElemsAssigned,FSideWall.CheckAssigned);
  if FBackWall <> Nil then 
    Inc(ElemsAssigned,FBackWall.CheckAssigned);
  if FPlotArea <> Nil then 
    Inc(ElemsAssigned,FPlotArea.CheckAssigned);
  if FLegend <> Nil then 
    Inc(ElemsAssigned,FLegend.CheckAssigned);
  if FPlotVisOnly <> Nil then 
    Inc(ElemsAssigned,FPlotVisOnly.CheckAssigned);
  if FDispBlanksAs <> Nil then 
    Inc(ElemsAssigned,FDispBlanksAs.CheckAssigned);
  if FShowDLblsOverMax <> Nil then 
    Inc(ElemsAssigned,FShowDLblsOverMax.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_Chart.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashB of
    $08DF99F7: begin
      if FTitle = Nil then 
        FTitle := TCT_Title.Create(FOwner);
      Result := FTitle;
    end;
    $5361A211: begin
      if FAutoTitleDeleted = Nil then 
        FAutoTitleDeleted := TCT_Boolean.Create(FOwner);
      Result := FAutoTitleDeleted;
    end;
    $A70B3497: begin
      if FPivotFmts = Nil then 
        FPivotFmts := TCT_PivotFmts.Create(FOwner);
      Result := FPivotFmts;
    end;
    $D386C983: begin
      if FView3D = Nil then 
        FView3D := TCT_View3D.Create(FOwner);
      Result := FView3D;
    end;
    $E30D225B: begin
      if FFloor = Nil then 
        FFloor := TCT_Surface.Create(FOwner);
      Result := FFloor;
    end;
    $EF329058: begin
      if FSideWall = Nil then 
        FSideWall := TCT_Surface.Create(FOwner);
      Result := FSideWall;
    end;
    $95B7666E: begin
      if FBackWall = Nil then 
        FBackWall := TCT_Surface.Create(FOwner);
      Result := FBackWall;
    end;
    $902C0893: begin
      if FPlotArea = Nil then 
        FPlotArea := TCT_PlotArea.Create(FOwner);
      Result := FPlotArea;
    end;
    $AF092E22: begin
      if FLegend = Nil then 
        FLegend := TCT_Legend.Create(FOwner);
      Result := FLegend;
    end;
    $C3216E86: begin
      if FPlotVisOnly = Nil then 
        FPlotVisOnly := TCT_Boolean.Create(FOwner);
      Result := FPlotVisOnly;
    end;
    $C14A1C94: begin
      if FDispBlanksAs = Nil then 
        FDispBlanksAs := TCT_DispBlanksAs.Create(FOwner);
      Result := FDispBlanksAs;
    end;
    $A9DC704B: begin
      if FShowDLblsOverMax = Nil then 
        FShowDLblsOverMax := TCT_Boolean.Create(FOwner);
      Result := FShowDLblsOverMax;
    end;
    $2EC15CAD: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_Chart.Write(AWriter: TXpgWriteXML);
begin
  if (FTitle <> Nil) and FTitle.Assigned then 
    if xaElements in FTitle.Assigneds then
    begin
      AWriter.BeginTag('c:title');
      FTitle.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:title');
  if (FAutoTitleDeleted <> Nil) and FAutoTitleDeleted.Assigned then 
  begin
    FAutoTitleDeleted.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:autoTitleDeleted');
  end;
  if (FPivotFmts <> Nil) and FPivotFmts.Assigned then 
    if xaElements in FPivotFmts.Assigneds then
    begin
      AWriter.BeginTag('c:pivotFmts');
      FPivotFmts.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:pivotFmts');
  if FView3D <> Nil then begin
    if FView3D.Assigned and (xaElements in FView3D.Assigneds) then
    begin
      AWriter.BeginTag('c:view3D');
      FView3D.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:view3D');
  end;
  if (FFloor <> Nil) and FFloor.Assigned then 
    if xaElements in FFloor.Assigneds then
    begin
      AWriter.BeginTag('c:floor');
      FFloor.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:floor');
  if (FSideWall <> Nil) and FSideWall.Assigned then 
    if xaElements in FSideWall.Assigneds then
    begin
      AWriter.BeginTag('c:sideWall');
      FSideWall.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:sideWall');
  if (FBackWall <> Nil) and FBackWall.Assigned then 
    if xaElements in FBackWall.Assigneds then
    begin
      AWriter.BeginTag('c:backWall');
      FBackWall.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:backWall');
  if (FPlotArea <> Nil) and FPlotArea.Assigned then 
    if xaElements in FPlotArea.Assigneds then
    begin
      AWriter.BeginTag('c:plotArea');
      FPlotArea.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:plotArea');

  if FLegend <> Nil then begin
    if FLegend.Assigned and (xaElements in FLegend.Assigneds) then
    begin
      AWriter.BeginTag('c:legend');
      FLegend.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:legend');
  end;

  if (FPlotVisOnly <> Nil) and FPlotVisOnly.Assigned then
  begin
    FPlotVisOnly.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:plotVisOnly');
  end;
  if (FDispBlanksAs <> Nil) and FDispBlanksAs.Assigned then 
  begin
    FDispBlanksAs.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:dispBlanksAs');
  end;
  if (FShowDLblsOverMax <> Nil) and FShowDLblsOverMax.Assigned then 
  begin
    FShowDLblsOverMax.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:showDLblsOverMax');
  end;
{$ifdef XLS_WRITE_EXTLST}
  if (FExtLst <> Nil) and FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then
    begin
      AWriter.BeginTag('c:extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:extLst');
{$endif}
end;

constructor TCT_Chart.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 13;
  FAttributeCount := 0;
end;

destructor TCT_Chart.Destroy;
begin
  if FTitle <> Nil then 
    FTitle.Free;
  if FAutoTitleDeleted <> Nil then 
    FAutoTitleDeleted.Free;
  if FPivotFmts <> Nil then 
    FPivotFmts.Free;
  if FView3D <> Nil then 
    FView3D.Free;
  if FFloor <> Nil then 
    FFloor.Free;
  if FSideWall <> Nil then 
    FSideWall.Free;
  if FBackWall <> Nil then 
    FBackWall.Free;
  if FPlotArea <> Nil then 
    FPlotArea.Free;
  if FLegend <> Nil then 
    FLegend.Free;
  if FPlotVisOnly <> Nil then 
    FPlotVisOnly.Free;
  if FDispBlanksAs <> Nil then 
    FDispBlanksAs.Free;
  if FShowDLblsOverMax <> Nil then 
    FShowDLblsOverMax.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_Chart.Clear;
begin
  FAssigneds := [];
  if FTitle <> Nil then 
    FreeAndNil(FTitle);
  if FAutoTitleDeleted <> Nil then 
    FreeAndNil(FAutoTitleDeleted);
  if FPivotFmts <> Nil then 
    FreeAndNil(FPivotFmts);
  if FView3D <> Nil then 
    FreeAndNil(FView3D);
  if FFloor <> Nil then 
    FreeAndNil(FFloor);
  if FSideWall <> Nil then 
    FreeAndNil(FSideWall);
  if FBackWall <> Nil then 
    FreeAndNil(FBackWall);
  if FPlotArea <> Nil then 
    FreeAndNil(FPlotArea);
  if FLegend <> Nil then 
    FreeAndNil(FLegend);
  if FPlotVisOnly <> Nil then 
    FreeAndNil(FPlotVisOnly);
  if FDispBlanksAs <> Nil then 
    FreeAndNil(FDispBlanksAs);
  if FShowDLblsOverMax <> Nil then 
    FreeAndNil(FShowDLblsOverMax);
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
end;

procedure TCT_Chart.Assign(AItem: TCT_Chart);
begin
end;

procedure TCT_Chart.CopyTo(AItem: TCT_Chart);
begin
end;

function  TCT_Chart.Create_Title: TCT_Title;
begin
  Result := TCT_Title.Create(FOwner);
  FTitle := Result;
end;

function  TCT_Chart.Create_AutoTitleDeleted: TCT_Boolean;
begin
  Result := TCT_Boolean.Create(FOwner);
  FAutoTitleDeleted := Result;
end;

function  TCT_Chart.Create_PivotFmts: TCT_PivotFmts;
begin
  Result := TCT_PivotFmts.Create(FOwner);
  FPivotFmts := Result;
end;

function  TCT_Chart.Create_View3D: TCT_View3D;
begin
  Result := TCT_View3D.Create(FOwner);
  FView3D := Result;
end;

function  TCT_Chart.Create_Floor: TCT_Surface;
begin
  Result := TCT_Surface.Create(FOwner);
  FFloor := Result;
end;

function  TCT_Chart.Create_SideWall: TCT_Surface;
begin
  Result := TCT_Surface.Create(FOwner);
  FSideWall := Result;
end;

function  TCT_Chart.Create_BackWall: TCT_Surface;
begin
  Result := TCT_Surface.Create(FOwner);
  FBackWall := Result;
end;

function  TCT_Chart.Create_PlotArea: TCT_PlotArea;
begin
  Result := TCT_PlotArea.Create(FOwner);
  FPlotArea := Result;
end;

function  TCT_Chart.Create_Legend: TCT_Legend;
begin
  Result := TCT_Legend.Create(FOwner);
  FLegend := Result;
end;

function  TCT_Chart.Create_PlotVisOnly: TCT_Boolean;
begin
  Result := TCT_Boolean.Create(FOwner);
  FPlotVisOnly := Result;
end;

function  TCT_Chart.Create_DispBlanksAs: TCT_DispBlanksAs;
begin
  Result := TCT_DispBlanksAs.Create(FOwner);
  FDispBlanksAs := Result;
end;

function  TCT_Chart.Create_ShowDLblsOverMax: TCT_Boolean;
begin
  Result := TCT_Boolean.Create(FOwner);
  FShowDLblsOverMax := Result;
end;

function  TCT_Chart.Create_ExtLst: TCT_ExtensionList;
begin
  Result := TCT_ExtensionList.Create(FOwner);
  FExtLst := Result;
end;

{ TCT_ExternalData }

function  TCT_ExternalData.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [xaAttributes];
  if FAutoUpdate <> Nil then
    Inc(ElemsAssigned,FAutoUpdate.CheckAssigned);
  Result := 1;
  if ElemsAssigned > 0 then
    FAssigneds := FAssigneds + [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_ExternalData.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'c:autoUpdate' then 
  begin
    if FAutoUpdate = Nil then 
      FAutoUpdate := TCT_Boolean.Create(FOwner);
    Result := FAutoUpdate;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_ExternalData.Write(AWriter: TXpgWriteXML);
begin
  if (FAutoUpdate <> Nil) and FAutoUpdate.Assigned then 
  begin
    FAutoUpdate.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:autoUpdate');
  end;
end;

procedure TCT_ExternalData.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('r:id',FR_Id);
end;

procedure TCT_ExternalData.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'r:id' then 
    FR_Id := AAttributes.Values[0]
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_ExternalData.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
end;

destructor TCT_ExternalData.Destroy;
begin
  if FAutoUpdate <> Nil then 
    FAutoUpdate.Free;
end;

procedure TCT_ExternalData.Clear;
begin
  FAssigneds := [];
  if FAutoUpdate <> Nil then 
    FreeAndNil(FAutoUpdate);
  FR_Id := '';
end;

procedure TCT_ExternalData.Assign(AItem: TCT_ExternalData);
begin
end;

procedure TCT_ExternalData.CopyTo(AItem: TCT_ExternalData);
begin
end;

function  TCT_ExternalData.Create_AutoUpdate: TCT_Boolean;
begin
  Result := TCT_Boolean.Create(FOwner);
  FAutoUpdate := Result;
end;

{ TCT_PrintSettings }

function  TCT_PrintSettings.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FHeaderFooter <> Nil then 
    Inc(ElemsAssigned,FHeaderFooter.CheckAssigned);
  if FPageMargins <> Nil then 
    Inc(ElemsAssigned,FPageMargins.CheckAssigned);
  if FPageSetup <> Nil then 
    Inc(ElemsAssigned,FPageSetup.CheckAssigned);
  if FLegacyDrawingHF <> Nil then 
    Inc(ElemsAssigned,FLegacyDrawingHF.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_PrintSettings.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $00000575: begin
      if FHeaderFooter = Nil then 
        FHeaderFooter := TCT_HeaderFooter.Create(FOwner);
      Result := FHeaderFooter;
    end;
    $0000050B: begin
      if FPageMargins = Nil then 
        FPageMargins := TCT_PageMargins.Create(FOwner);
      Result := FPageMargins;
    end;
    $0000044B: begin
      if FPageSetup = Nil then 
        FPageSetup := TCT_PageSetup.Create(FOwner);
      Result := FPageSetup;
    end;
    $0000066C: begin
      if FLegacyDrawingHF = Nil then 
        FLegacyDrawingHF := TCT_RelId.Create(FOwner);
      Result := FLegacyDrawingHF;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_PrintSettings.Write(AWriter: TXpgWriteXML);
begin
  if (FHeaderFooter <> Nil) and FHeaderFooter.Assigned then 
  begin
    FHeaderFooter.WriteAttributes(AWriter);
    if xaElements in FHeaderFooter.Assigneds then
    begin
      AWriter.BeginTag('c:headerFooter');
      FHeaderFooter.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:headerFooter');
  end;
  if (FPageMargins <> Nil) and FPageMargins.Assigned then 
  begin
    FPageMargins.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:pageMargins');
  end;
  if (FPageSetup <> Nil) and FPageSetup.Assigned then 
  begin
    FPageSetup.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:pageSetup');
  end;
  if (FLegacyDrawingHF <> Nil) and FLegacyDrawingHF.Assigned then 
  begin
    FLegacyDrawingHF.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:legacyDrawingHF');
  end;
end;

constructor TCT_PrintSettings.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 4;
  FAttributeCount := 0;
end;

destructor TCT_PrintSettings.Destroy;
begin
  if FHeaderFooter <> Nil then 
    FHeaderFooter.Free;
  if FPageMargins <> Nil then 
    FPageMargins.Free;
  if FPageSetup <> Nil then 
    FPageSetup.Free;
  if FLegacyDrawingHF <> Nil then 
    FLegacyDrawingHF.Free;
end;

procedure TCT_PrintSettings.Clear;
begin
  FAssigneds := [];
  if FHeaderFooter <> Nil then 
    FreeAndNil(FHeaderFooter);
  if FPageMargins <> Nil then 
    FreeAndNil(FPageMargins);
  if FPageSetup <> Nil then 
    FreeAndNil(FPageSetup);
  if FLegacyDrawingHF <> Nil then 
    FreeAndNil(FLegacyDrawingHF);
end;

procedure TCT_PrintSettings.Assign(AItem: TCT_PrintSettings);
begin
end;

procedure TCT_PrintSettings.CopyTo(AItem: TCT_PrintSettings);
begin
end;

function  TCT_PrintSettings.Create_HeaderFooter: TCT_HeaderFooter;
begin
  Result := TCT_HeaderFooter.Create(FOwner);
  FHeaderFooter := Result;
end;

function  TCT_PrintSettings.Create_PageMargins: TCT_PageMargins;
begin
  Result := TCT_PageMargins.Create(FOwner);
  FPageMargins := Result;
end;

function  TCT_PrintSettings.Create_PageSetup: TCT_PageSetup;
begin
  Result := TCT_PageSetup.Create(FOwner);
  FPageSetup := Result;
end;

function  TCT_PrintSettings.Create_LegacyDrawingHF: TCT_RelId;
begin
  Result := TCT_RelId.Create(FOwner);
  FLegacyDrawingHF := Result;
end;

{ TCT_EmptyElement }

function  TCT_EmptyElement.CheckAssigned: integer;
begin
  FAssigneds := [];
  Result := 0;
end;

procedure TCT_EmptyElement.Write(AWriter: TXpgWriteXML);
begin
end;

constructor TCT_EmptyElement.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 0;
end;

destructor TCT_EmptyElement.Destroy;
begin
end;

procedure TCT_EmptyElement.Clear;
begin
  FAssigneds := [];
end;

procedure TCT_EmptyElement.Assign(AItem: TCT_EmptyElement);
begin
end;

procedure TCT_EmptyElement.CopyTo(AItem: TCT_EmptyElement);
begin
end;

{ TCT_BaseStylesOverride }

function  TCT_BaseStylesOverride.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FA_ClrScheme <> Nil then 
    Inc(ElemsAssigned,FA_ClrScheme.CheckAssigned);
  if FA_FontScheme <> Nil then 
    Inc(ElemsAssigned,FA_FontScheme.CheckAssigned);
  if FA_FmtScheme <> Nil then 
    Inc(ElemsAssigned,FA_FmtScheme.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_BaseStylesOverride.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $00000431: begin
      if FA_ClrScheme = Nil then 
        FA_ClrScheme := TCT_ColorScheme.Create(FOwner);
      Result := FA_ClrScheme;
    end;
    $000004A7: begin
      if FA_FontScheme = Nil then 
        FA_FontScheme := TCT_FontScheme.Create(FOwner);
      Result := FA_FontScheme;
    end;
    $00000437: begin
      if FA_FmtScheme = Nil then 
        FA_FmtScheme := TCT_StyleMatrix.Create(FOwner);
      Result := FA_FmtScheme;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_BaseStylesOverride.Write(AWriter: TXpgWriteXML);
begin
  if (FA_ClrScheme <> Nil) and FA_ClrScheme.Assigned then 
  begin
    FA_ClrScheme.WriteAttributes(AWriter);
    if xaElements in FA_ClrScheme.Assigneds then
    begin
      AWriter.BeginTag('a:clrScheme');
      FA_ClrScheme.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:clrScheme');
  end;
  if (FA_FontScheme <> Nil) and FA_FontScheme.Assigned then 
  begin
    FA_FontScheme.WriteAttributes(AWriter);
    if xaElements in FA_FontScheme.Assigneds then
    begin
      AWriter.BeginTag('a:fontScheme');
      FA_FontScheme.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:fontScheme');
  end;
  if (FA_FmtScheme <> Nil) and FA_FmtScheme.Assigned then 
  begin
    FA_FmtScheme.WriteAttributes(AWriter);
    if xaElements in FA_FmtScheme.Assigneds then
    begin
      AWriter.BeginTag('a:fmtScheme');
      FA_FmtScheme.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:fmtScheme');
  end;
end;

constructor TCT_BaseStylesOverride.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 3;
  FAttributeCount := 0;
end;

destructor TCT_BaseStylesOverride.Destroy;
begin
  if FA_ClrScheme <> Nil then 
    FA_ClrScheme.Free;
  if FA_FontScheme <> Nil then 
    FA_FontScheme.Free;
  if FA_FmtScheme <> Nil then 
    FA_FmtScheme.Free;
end;

procedure TCT_BaseStylesOverride.Clear;
begin
  FAssigneds := [];
  if FA_ClrScheme <> Nil then 
    FreeAndNil(FA_ClrScheme);
  if FA_FontScheme <> Nil then 
    FreeAndNil(FA_FontScheme);
  if FA_FmtScheme <> Nil then 
    FreeAndNil(FA_FmtScheme);
end;

procedure TCT_BaseStylesOverride.Assign(AItem: TCT_BaseStylesOverride);
begin
end;

procedure TCT_BaseStylesOverride.CopyTo(AItem: TCT_BaseStylesOverride);
begin
end;

function  TCT_BaseStylesOverride.Create_A_ClrScheme: TCT_ColorScheme;
begin
  Result := TCT_ColorScheme.Create(FOwner);
  FA_ClrScheme := Result;
end;

function  TCT_BaseStylesOverride.Create_A_FontScheme: TCT_FontScheme;
begin
  Result := TCT_FontScheme.Create(FOwner);
  FA_FontScheme := Result;
end;

function  TCT_BaseStylesOverride.Create_A_FmtScheme: TCT_StyleMatrix;
begin
  Result := TCT_StyleMatrix.Create(FOwner);
  FA_FmtScheme := Result;
end;

{ TCT_OfficeStyleSheet }

function  TCT_OfficeStyleSheet.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FName <> '' then 
    Inc(AttrsAssigned);
  if FA_ThemeElements <> Nil then 
    Inc(ElemsAssigned,FA_ThemeElements.CheckAssigned);
  if FA_ObjectDefaults <> Nil then 
    Inc(ElemsAssigned,FA_ObjectDefaults.CheckAssigned);
  if FA_ExtraClrSchemeLst <> Nil then 
    Inc(ElemsAssigned,FA_ExtraClrSchemeLst.CheckAssigned);
  if FA_CustClrLst <> Nil then 
    Inc(ElemsAssigned,FA_CustClrLst.CheckAssigned);
  if FA_ExtLst <> Nil then 
    Inc(ElemsAssigned,FA_ExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_OfficeStyleSheet.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $000005EB: begin
      if FA_ThemeElements = Nil then 
        FA_ThemeElements := TCT_BaseStyles.Create(FOwner);
      Result := FA_ThemeElements;
    end;
    $0000064A: begin
      if FA_ObjectDefaults = Nil then 
        FA_ObjectDefaults := TCT_ObjectStyleDefaults.Create(FOwner);
      Result := FA_ObjectDefaults;
    end;
    $00000768: begin
      if FA_ExtraClrSchemeLst = Nil then 
        FA_ExtraClrSchemeLst := TCT_ColorSchemeList.Create(FOwner);
      Result := FA_ExtraClrSchemeLst;
    end;
    $000004AE: begin
      if FA_CustClrLst = Nil then 
        FA_CustClrLst := TCT_CustomColorList.Create(FOwner);
      Result := FA_CustClrLst;
    end;
    $0000031F: begin
      if FA_ExtLst = Nil then 
        FA_ExtLst := TCT_OfficeArtExtensionList.Create(FOwner);
      Result := FA_ExtLst;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_OfficeStyleSheet.Write(AWriter: TXpgWriteXML);
begin
  if (FA_ThemeElements <> Nil) and FA_ThemeElements.Assigned then 
    if xaElements in FA_ThemeElements.Assigneds then
    begin
      AWriter.BeginTag('a:themeElements');
      FA_ThemeElements.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:themeElements')
  else 
    AWriter.SimpleTag('a:themeElements');
  if (FA_ObjectDefaults <> Nil) and FA_ObjectDefaults.Assigned then 
    if xaElements in FA_ObjectDefaults.Assigneds then
    begin
      AWriter.BeginTag('a:objectDefaults');
      FA_ObjectDefaults.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:objectDefaults');
  if (FA_ExtraClrSchemeLst <> Nil) and FA_ExtraClrSchemeLst.Assigned then 
    if xaElements in FA_ExtraClrSchemeLst.Assigneds then
    begin
      AWriter.BeginTag('a:extraClrSchemeLst');
      FA_ExtraClrSchemeLst.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:extraClrSchemeLst');
  if (FA_CustClrLst <> Nil) and FA_CustClrLst.Assigned then 
    if xaElements in FA_CustClrLst.Assigneds then
    begin
      AWriter.BeginTag('a:custClrLst');
      FA_CustClrLst.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:custClrLst');
  if (FA_ExtLst <> Nil) and FA_ExtLst.Assigned then 
    if xaElements in FA_ExtLst.Assigneds then
    begin
      AWriter.BeginTag('a:extLst');
      FA_ExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:extLst');
end;

procedure TCT_OfficeStyleSheet.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FName <> '' then 
    AWriter.AddAttribute('name',FName);
end;

procedure TCT_OfficeStyleSheet.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'name' then 
    FName := AAttributes.Values[0]
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_OfficeStyleSheet.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 5;
  FAttributeCount := 1;
end;

destructor TCT_OfficeStyleSheet.Destroy;
begin
  if FA_ThemeElements <> Nil then 
    FA_ThemeElements.Free;
  if FA_ObjectDefaults <> Nil then 
    FA_ObjectDefaults.Free;
  if FA_ExtraClrSchemeLst <> Nil then 
    FA_ExtraClrSchemeLst.Free;
  if FA_CustClrLst <> Nil then 
    FA_CustClrLst.Free;
  if FA_ExtLst <> Nil then 
    FA_ExtLst.Free;
end;

procedure TCT_OfficeStyleSheet.Clear;
begin
  FAssigneds := [];
  if FA_ThemeElements <> Nil then 
    FreeAndNil(FA_ThemeElements);
  if FA_ObjectDefaults <> Nil then 
    FreeAndNil(FA_ObjectDefaults);
  if FA_ExtraClrSchemeLst <> Nil then 
    FreeAndNil(FA_ExtraClrSchemeLst);
  if FA_CustClrLst <> Nil then 
    FreeAndNil(FA_CustClrLst);
  if FA_ExtLst <> Nil then 
    FreeAndNil(FA_ExtLst);
  FName := '';
end;

procedure TCT_OfficeStyleSheet.Assign(AItem: TCT_OfficeStyleSheet);
begin
end;

procedure TCT_OfficeStyleSheet.CopyTo(AItem: TCT_OfficeStyleSheet);
begin
end;

function  TCT_OfficeStyleSheet.Create_A_ThemeElements: TCT_BaseStyles;
begin
  Result := TCT_BaseStyles.Create(FOwner);
  FA_ThemeElements := Result;
end;

function  TCT_OfficeStyleSheet.Create_A_ObjectDefaults: TCT_ObjectStyleDefaults;
begin
  Result := TCT_ObjectStyleDefaults.Create(FOwner);
  FA_ObjectDefaults := Result;
end;

function  TCT_OfficeStyleSheet.Create_A_ExtraClrSchemeLst: TCT_ColorSchemeList;
begin
  Result := TCT_ColorSchemeList.Create(FOwner);
  FA_ExtraClrSchemeLst := Result;
end;

function  TCT_OfficeStyleSheet.Create_A_CustClrLst: TCT_CustomColorList;
begin
  Result := TCT_CustomColorList.Create(FOwner);
  FA_CustClrLst := Result;
end;

function  TCT_OfficeStyleSheet.Create_A_ExtLst: TCT_OfficeArtExtensionList;
begin
  Result := TCT_OfficeArtExtensionList.Create(FOwner);
  FA_ExtLst := Result;
end;

{ TCT_ChartSpace }

function  TCT_ChartSpace.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FDate1904 <> Nil then 
    Inc(ElemsAssigned,FDate1904.CheckAssigned);
  if FLang <> Nil then 
    Inc(ElemsAssigned,FLang.CheckAssigned);
  if FRoundedCorners <> Nil then 
    Inc(ElemsAssigned,FRoundedCorners.CheckAssigned);
  if FStyle <> Nil then 
    Inc(ElemsAssigned,FStyle.CheckAssigned);
  if FClrMapOvr <> Nil then 
    Inc(ElemsAssigned,FClrMapOvr.CheckAssigned);
  if FPivotSource <> Nil then 
    Inc(ElemsAssigned,FPivotSource.CheckAssigned);
  if FProtection <> Nil then 
    Inc(ElemsAssigned,FProtection.CheckAssigned);
  if FChart <> Nil then 
    Inc(ElemsAssigned,FChart.CheckAssigned);
  if FSpPr <> Nil then 
    Inc(ElemsAssigned,FSpPr.CheckAssigned);
  if FTxPr <> Nil then 
    Inc(ElemsAssigned,FTxPr.CheckAssigned);
  if FExternalData <> Nil then 
    Inc(ElemsAssigned,FExternalData.CheckAssigned);
  if FPrintSettings <> Nil then 
    Inc(ElemsAssigned,FPrintSettings.CheckAssigned);
  if FUserShapes <> Nil then 
    Inc(ElemsAssigned,FUserShapes.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_ChartSpace.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $00000309: begin
      if FDate1904 = Nil then 
        FDate1904 := TCT_Boolean.Create(FOwner);
      Result := FDate1904;
    end;
    $0000023F: begin
      if FLang = Nil then 
        FLang := TCT_TextLanguageID.Create(FOwner);
      Result := FLang;
    end;
    $0000066A: begin
      if FRoundedCorners = Nil then 
        FRoundedCorners := TCT_Boolean.Create(FOwner);
      Result := FRoundedCorners;
    end;
    $000002CE: begin
      if FStyle = Nil then 
        FStyle := TCT_Style.Create(FOwner);
      Result := FStyle;
    end;
    $00000433: begin
      if FClrMapOvr = Nil then 
        FClrMapOvr := TCT_ColorMapping.Create(FOwner);
      Result := FClrMapOvr;
    end;
    $00000540: begin
      if FPivotSource = Nil then 
        FPivotSource := TCT_PivotSource.Create(FOwner);
      Result := FPivotSource;
    end;
    $000004E4: begin
      if FProtection = Nil then 
        FProtection := TCT_Protection.Create(FOwner);
      Result := FProtection;
    end;
    $000002AF: begin
      if FChart = Nil then 
        FChart := TCT_Chart.Create(FOwner);
      Result := FChart;
    end;
    $00000242: begin
      if FSpPr = Nil then 
        FSpPr := TCT_ShapeProperties.Create(FOwner);
      Result := FSpPr;
    end;
    $0000024B: begin
      if FTxPr = Nil then 
        FTxPr := TCT_TextBody.Create(FOwner);
      Result := FTxPr;
    end;
    $0000057A: begin
      if FExternalData = Nil then 
        FExternalData := TCT_ExternalData.Create(FOwner);
      Result := FExternalData;
    end;
    $0000061B: begin
      if FPrintSettings = Nil then 
        FPrintSettings := TCT_PrintSettings.Create(FOwner);
      Result := FPrintSettings;
    end;
    $000004C0: begin
      if FUserShapes = Nil then 
        FUserShapes := TCT_RelId.Create(FOwner);
      Result := FUserShapes;
    end;
    $00000321: begin
      if FExtLst = Nil then
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else begin
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
    end;
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_ChartSpace.Write(AWriter: TXpgWriteXML);
begin
  if (FDate1904 <> Nil) and FDate1904.Assigned then
  begin
    FDate1904.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:date1904');
  end;
  if (FLang <> Nil) and FLang.Assigned then
  begin
    FLang.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:lang');
  end;
  if (FRoundedCorners <> Nil) and FRoundedCorners.Assigned then
  begin
    FRoundedCorners.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:roundedCorners');
  end;
  if (FStyle <> Nil) and FStyle.Assigned then 
  begin
    FStyle.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:style');
  end;
  if (FClrMapOvr <> Nil) and FClrMapOvr.Assigned then 
  begin
    FClrMapOvr.WriteAttributes(AWriter);
    if xaElements in FClrMapOvr.Assigneds then
    begin
      AWriter.BeginTag('c:clrMapOvr');
      FClrMapOvr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:clrMapOvr');
  end;
  if (FPivotSource <> Nil) and FPivotSource.Assigned then 
    if xaElements in FPivotSource.Assigneds then
    begin
      AWriter.BeginTag('c:pivotSource');
      FPivotSource.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:pivotSource');
  if (FProtection <> Nil) and FProtection.Assigned then 
    if xaElements in FProtection.Assigneds then
    begin
      AWriter.BeginTag('c:protection');
      FProtection.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:protection');
  if (FChart <> Nil) and FChart.Assigned then 
    if xaElements in FChart.Assigneds then
    begin
      AWriter.BeginTag('c:chart');
      FChart.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:chart');
  if (FSpPr <> Nil) and FSpPr.Assigned then 
  begin
    FSpPr.WriteAttributes(AWriter);
    if xaElements in FSpPr.Assigneds then
    begin
      AWriter.BeginTag('c:spPr');
      FSpPr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('c:spPr');
  end;
  if (FTxPr <> Nil) and FTxPr.Assigned then
    if xaElements in FTxPr.Assigneds then
    begin
      AWriter.BeginTag('c:txPr');
      FTxPr.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:txPr');
  if (FExternalData <> Nil) and FExternalData.Assigned then
  begin
    FExternalData.WriteAttributes(AWriter);
    if xaElements in FExternalData.Assigneds then
    begin
      AWriter.BeginTag('c:externalData');
      FExternalData.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:externalData');
  end;
  if (FPrintSettings <> Nil) and FPrintSettings.Assigned then
    if xaElements in FPrintSettings.Assigneds then
    begin
      AWriter.BeginTag('c:printSettings');
      FPrintSettings.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:printSettings');
  if (FUserShapes <> Nil) and FUserShapes.Assigned then
  begin
    FUserShapes.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:userShapes');
  end;
{$ifdef XLS_WRITE_EXTLST}
  if (FExtLst <> Nil) and FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then
    begin
      AWriter.BeginTag('c:extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:extLst');
{$endif}
end;

constructor TCT_ChartSpace.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 14;
  FAttributeCount := 0;
end;

destructor TCT_ChartSpace.Destroy;
begin
  if FDate1904 <> Nil then 
    FDate1904.Free;
  if FLang <> Nil then 
    FLang.Free;
  if FRoundedCorners <> Nil then 
    FRoundedCorners.Free;
  if FStyle <> Nil then 
    FStyle.Free;
  if FClrMapOvr <> Nil then 
    FClrMapOvr.Free;
  if FPivotSource <> Nil then 
    FPivotSource.Free;
  if FProtection <> Nil then 
    FProtection.Free;
  if FChart <> Nil then 
    FChart.Free;
  if FSpPr <> Nil then 
    FSpPr.Free;
  if FTxPr <> Nil then 
    FTxPr.Free;
  if FExternalData <> Nil then 
    FExternalData.Free;
  if FPrintSettings <> Nil then 
    FPrintSettings.Free;
  if FUserShapes <> Nil then 
    FUserShapes.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_ChartSpace.Clear;
begin
  FAssigneds := [];
  if FDate1904 <> Nil then 
    FreeAndNil(FDate1904);
  if FLang <> Nil then 
    FreeAndNil(FLang);
  if FRoundedCorners <> Nil then 
    FreeAndNil(FRoundedCorners);
  if FStyle <> Nil then 
    FreeAndNil(FStyle);
  if FClrMapOvr <> Nil then 
    FreeAndNil(FClrMapOvr);
  if FPivotSource <> Nil then 
    FreeAndNil(FPivotSource);
  if FProtection <> Nil then 
    FreeAndNil(FProtection);
  if FChart <> Nil then 
    FreeAndNil(FChart);
  if FSpPr <> Nil then 
    FreeAndNil(FSpPr);
  if FTxPr <> Nil then 
    FreeAndNil(FTxPr);
  if FExternalData <> Nil then 
    FreeAndNil(FExternalData);
  if FPrintSettings <> Nil then 
    FreeAndNil(FPrintSettings);
  if FUserShapes <> Nil then 
    FreeAndNil(FUserShapes);
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
end;

procedure TCT_ChartSpace.Assign(AItem: TCT_ChartSpace);
begin
end;

procedure TCT_ChartSpace.CopyTo(AItem: TCT_ChartSpace);
begin
end;

function  TCT_ChartSpace.Create_Date1904: TCT_Boolean;
begin
  Result := TCT_Boolean.Create(FOwner);
  FDate1904 := Result;
end;

function  TCT_ChartSpace.Create_Lang: TCT_TextLanguageID;
begin
  Result := TCT_TextLanguageID.Create(FOwner);
  FLang := Result;
end;

function  TCT_ChartSpace.Create_RoundedCorners: TCT_Boolean;
begin
  Result := TCT_Boolean.Create(FOwner);
  FRoundedCorners := Result;
end;

function  TCT_ChartSpace.Create_Style: TCT_Style;
begin
  Result := TCT_Style.Create(FOwner);
  FStyle := Result;
end;

function  TCT_ChartSpace.Create_ClrMapOvr: TCT_ColorMapping;
begin
  Result := TCT_ColorMapping.Create(FOwner);
  FClrMapOvr := Result;
end;

function  TCT_ChartSpace.Create_PivotSource: TCT_PivotSource;
begin
  Result := TCT_PivotSource.Create(FOwner);
  FPivotSource := Result;
end;

function  TCT_ChartSpace.Create_Protection: TCT_Protection;
begin
  Result := TCT_Protection.Create(FOwner);
  FProtection := Result;
end;

function  TCT_ChartSpace.Create_Chart: TCT_Chart;
begin
  Result := TCT_Chart.Create(FOwner);
  FChart := Result;
end;

function  TCT_ChartSpace.Create_SpPr: TCT_ShapeProperties;
begin
  Result := TCT_ShapeProperties.Create(FOwner);
  FSpPr := Result;
end;

function  TCT_ChartSpace.Create_TxPr: TCT_TextBody;
begin
  Result := TCT_TextBody.Create(FOwner);
  FTxPr := Result;
end;

function  TCT_ChartSpace.Create_ExternalData: TCT_ExternalData;
begin
  Result := TCT_ExternalData.Create(FOwner);
  FExternalData := Result;
end;

function  TCT_ChartSpace.Create_PrintSettings: TCT_PrintSettings;
begin
  Result := TCT_PrintSettings.Create(FOwner);
  FPrintSettings := Result;
end;

function  TCT_ChartSpace.Create_UserShapes: TCT_RelId;
begin
  Result := TCT_RelId.Create(FOwner);
  FUserShapes := Result;
end;

function  TCT_ChartSpace.Create_ExtLst: TCT_ExtensionList;
begin
  Result := TCT_ExtensionList.Create(FOwner);
  FExtLst := Result;
end;

{ TCT_ColorMappingOverride }

function  TCT_ColorMappingOverride.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FA_MasterClrMapping <> Nil then 
    Inc(ElemsAssigned,FA_MasterClrMapping.CheckAssigned);
  if FA_OverrideClrMapping <> Nil then 
    Inc(ElemsAssigned,FA_OverrideClrMapping.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_ColorMappingOverride.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $00000714: begin
      if FA_MasterClrMapping = Nil then 
        FA_MasterClrMapping := TCT_EmptyElement.Create(FOwner);
      Result := FA_MasterClrMapping;
    end;
    $000007E8: begin
      if FA_OverrideClrMapping = Nil then 
        FA_OverrideClrMapping := TCT_ColorMapping.Create(FOwner);
      Result := FA_OverrideClrMapping;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_ColorMappingOverride.Write(AWriter: TXpgWriteXML);
begin
  if (FA_MasterClrMapping <> Nil) and FA_MasterClrMapping.Assigned then 
    AWriter.SimpleTag('a:masterClrMapping')
  else 
    AWriter.SimpleTag('a:masterClrMapping');
  if (FA_OverrideClrMapping <> Nil) and FA_OverrideClrMapping.Assigned then 
  begin
    FA_OverrideClrMapping.WriteAttributes(AWriter);
    if xaElements in FA_OverrideClrMapping.Assigneds then
    begin
      AWriter.BeginTag('a:overrideClrMapping');
      FA_OverrideClrMapping.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:overrideClrMapping');
  end
  else 
    AWriter.SimpleTag('a:overrideClrMapping');
end;

constructor TCT_ColorMappingOverride.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 0;
end;

destructor TCT_ColorMappingOverride.Destroy;
begin
  if FA_MasterClrMapping <> Nil then 
    FA_MasterClrMapping.Free;
  if FA_OverrideClrMapping <> Nil then 
    FA_OverrideClrMapping.Free;
end;

procedure TCT_ColorMappingOverride.Clear;
begin
  FAssigneds := [];
  if FA_MasterClrMapping <> Nil then 
    FreeAndNil(FA_MasterClrMapping);
  if FA_OverrideClrMapping <> Nil then 
    FreeAndNil(FA_OverrideClrMapping);
end;

procedure TCT_ColorMappingOverride.Assign(AItem: TCT_ColorMappingOverride);
begin
end;

procedure TCT_ColorMappingOverride.CopyTo(AItem: TCT_ColorMappingOverride);
begin
end;

function  TCT_ColorMappingOverride.Create_A_MasterClrMapping: TCT_EmptyElement;
begin
  Result := TCT_EmptyElement.Create(FOwner);
  FA_MasterClrMapping := Result;
end;

function  TCT_ColorMappingOverride.Create_A_OverrideClrMapping: TCT_ColorMapping;
begin
  Result := TCT_ColorMapping.Create(FOwner);
  FA_OverrideClrMapping := Result;
end;

{ TCT_ClipboardStyleSheet }

function  TCT_ClipboardStyleSheet.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FA_ThemeElements <> Nil then 
    Inc(ElemsAssigned,FA_ThemeElements.CheckAssigned);
  if FA_ClrMap <> Nil then 
    Inc(ElemsAssigned,FA_ClrMap.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_ClipboardStyleSheet.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $000005EB: begin
      if FA_ThemeElements = Nil then 
        FA_ThemeElements := TCT_BaseStyles.Create(FOwner);
      Result := FA_ThemeElements;
    end;
    $000002FA: begin
      if FA_ClrMap = Nil then 
        FA_ClrMap := TCT_ColorMapping.Create(FOwner);
      Result := FA_ClrMap;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_ClipboardStyleSheet.Write(AWriter: TXpgWriteXML);
begin
  if (FA_ThemeElements <> Nil) and FA_ThemeElements.Assigned then 
    if xaElements in FA_ThemeElements.Assigneds then
    begin
      AWriter.BeginTag('a:themeElements');
      FA_ThemeElements.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:themeElements')
  else 
    AWriter.SimpleTag('a:themeElements');
  if (FA_ClrMap <> Nil) and FA_ClrMap.Assigned then 
  begin
    FA_ClrMap.WriteAttributes(AWriter);
    if xaElements in FA_ClrMap.Assigneds then
    begin
      AWriter.BeginTag('a:clrMap');
      FA_ClrMap.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:clrMap');
  end
  else 
    AWriter.SimpleTag('a:clrMap');
end;

constructor TCT_ClipboardStyleSheet.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 0;
end;

destructor TCT_ClipboardStyleSheet.Destroy;
begin
  if FA_ThemeElements <> Nil then 
    FA_ThemeElements.Free;
  if FA_ClrMap <> Nil then 
    FA_ClrMap.Free;
end;

procedure TCT_ClipboardStyleSheet.Clear;
begin
  FAssigneds := [];
  if FA_ThemeElements <> Nil then 
    FreeAndNil(FA_ThemeElements);
  if FA_ClrMap <> Nil then 
    FreeAndNil(FA_ClrMap);
end;

procedure TCT_ClipboardStyleSheet.Assign(AItem: TCT_ClipboardStyleSheet);
begin
end;

procedure TCT_ClipboardStyleSheet.CopyTo(AItem: TCT_ClipboardStyleSheet);
begin
end;

function  TCT_ClipboardStyleSheet.Create_A_ThemeElements: TCT_BaseStyles;
begin
  Result := TCT_BaseStyles.Create(FOwner);
  FA_ThemeElements := Result;
end;

function  TCT_ClipboardStyleSheet.Create_A_ClrMap: TCT_ColorMapping;
begin
  Result := TCT_ColorMapping.Create(FOwner);
  FA_ClrMap := Result;
end;

{ T__ROOT__ }

function  T__ROOT__.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FChart <> Nil then
    Inc(ElemsAssigned,FChart.CheckAssigned);
  if FChartSpace <> Nil then
    Inc(ElemsAssigned,FChartSpace.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  T__ROOT__.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  i: integer;
begin
  for i := 0 to AReader.Attributes.Count - 1 do
    FRootAttributes.Add(AReader.Attributes.AsXmlText2(i));
  Result := Self;
  case AReader.QNameHashA of
    $000002AF: begin
      if FChart = Nil then
        FChart := TCT_RelId.Create(FOwner);
      Result := FChart;
    end;
    $0000049B: begin
      if FChartSpace = Nil then
        FChartSpace := TCT_ChartSpace.Create(FOwner);
      Result := FChartSpace;
    end;
    else
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then
    Result.Assigneds := [xaRead];
end;

procedure T__ROOT__.Write(AWriter: TXpgWriteXML);
begin
  AWriter.Attributes := FRootAttributes.Text;
  if (FChart <> Nil) and FChart.Assigned then
  begin
    FChart.WriteAttributes(AWriter);
    AWriter.SimpleTag('c:chart');
  end;
  if (FChartSpace <> Nil) and FChartSpace.Assigned then begin
    if xaElements in FChartSpace.FAssigneds then
    begin
      AWriter.BeginTag('c:chartSpace');
      FChartSpace.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('c:chartSpace');
  end;
end;

constructor T__ROOT__.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FRootAttributes := TStringXpgList.Create;
  FElementCount := 6;
  FAttributeCount := 0;
end;

destructor T__ROOT__.Destroy;
begin
  FRootAttributes.Free;
  if FChart <> Nil then
    FChart.Free;
  if FChartSpace <> Nil then
    FChartSpace.Free;
end;

procedure T__ROOT__.Clear;
begin
  FRootAttributes.Clear;
  FAssigneds := [];
  if FChart <> Nil then
    FreeAndNil(FChart);
  if FChartSpace <> Nil then
    FreeAndNil(FChartSpace);
end;

procedure T__ROOT__.Assign(AItem: T__ROOT__);
begin
end;

procedure T__ROOT__.CopyTo(AItem: T__ROOT__);
begin
end;

function  T__ROOT__.Create_Chart: TCT_RelId;
begin
  Result := TCT_RelId.Create(FOwner);
  FChart := Result;
end;

function  T__ROOT__.Create_ChartSpace: TCT_ChartSpace;
begin
  Result := TCT_ChartSpace.Create(FOwner);
  FChartSpace := Result;
end;

{ TXPGDocXLSXChart }

function  TXPGDocXLSXChart.GetChart: TCT_RelId;
begin
  Result := FRoot.Chart;
end;

function  TXPGDocXLSXChart.GetChartSpace: TCT_ChartSpace;
begin
  Result := FRoot.ChartSpace;
end;

procedure TXPGDocXLSXChart.AddColors;
begin
  FColors := TXpgSimpleDOM.Create;
end;

procedure TXPGDocXLSXChart.AddStyle;
begin
  FStyle := TXpgSimpleDOM.Create;
end;

constructor TXPGDocXLSXChart.Create;
begin
  FRoot := T__ROOT__.Create(Self);
  FErrors := TXpgPErrors.Create;
  FReader := TXPGReader.Create(FErrors,FRoot);
  FWriter := TXpgWriteXML.Create;
end;

destructor TXPGDocXLSXChart.Destroy;
begin
  FRoot.Free;
  FReader.Free;
  FWriter.Free;
  FErrors.Free;

  if FStyle <> Nil then
    FStyle.Free;
  if FColors <> Nil then
    FColors.Free;
  inherited Destroy;
end;

procedure TXPGDocXLSXChart.LoadFromFile(AFilename: AxUCString);
var
  Stream: TFileStream;
begin
  Stream := TFileStream.Create(AFilename,fmOpenRead);
  try
    FReader.LoadFromStream(Stream);
  finally
    Stream.Free;
  end;
end;

procedure TXPGDocXLSXChart.LoadFromStream(AStream: TStream);
begin
  FReader.LoadFromStream(AStream);
end;

procedure TXPGDocXLSXChart.SaveToFile(AFilename: AxUCString; AClassToWrite: TClass);
begin
  FRoot.FCurrWriteClass := AClassToWrite;
  FWriter.SaveToFile(AFilename);
  FRoot.CheckAssigned;
  FRoot.Write(FWriter);
end;

procedure TXPGDocXLSXChart.SaveToStream(AStream: TStream);
begin
  FWriter.SaveToStream(AStream);
  FRoot.CheckAssigned;
  FRoot.Write(FWriter);
end;

{ TCT_Shape_Chart }

procedure TCT_Shape_Chart.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then
    FVal := TST_Shape(StrToEnum('sts' + AAttributes.Values[0]))
  else
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

function TCT_Shape_Chart.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FVal <> stsBox then
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then
    FAssigneds := [xaAttributes];
end;

procedure TCT_Shape_Chart.Clear;
begin
  FAssigneds := [];
  FVal := stsBox;
end;

constructor TCT_Shape_Chart.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := stsBox;
end;

destructor TCT_Shape_Chart.Destroy;
begin
  inherited;
end;

procedure TCT_Shape_Chart.Write(AWriter: TXpgWriteXML);
begin

end;

procedure TCT_Shape_Chart.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FVal <> stsBox then
    AWriter.AddAttribute('val',StrTST_Shape[Integer(FVal)]);
end;

end.
