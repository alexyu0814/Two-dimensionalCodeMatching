unit XLSClassFactory5;

interface

uses Classes, SysUtils;

type TXLSClassFactoryType = (xcftNames,xcftNamesMember,
                             xcftMergedCells,xcftMergedCellsMember,
                             xcftHyperlinks,xcftHyperlinksMember,
                             xcftDataValidations,xcftDataValidationsMember,
                             xcftConditionalFormat,xcftConditionalFormats,
                             xcftAutofilter,xcftAutofilterColumn,
                             xcftDrawing);

type TXLSClassFactory = class(TObject)
protected
public
     function CreateAClass(AClassType: TXLSClassFactoryType; AOwner: TObject = Nil): TObject; virtual; abstract;
//     function GetAClass(AClassType: TXLSClassFactoryType): TObject; virtual; abstract;
     end;

implementation

end.
