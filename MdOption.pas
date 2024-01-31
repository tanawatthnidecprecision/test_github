////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////
//////////////////////////////// Nidec-COPAL(Thailand) CO.,LTD                //
//////////////////////////////// Option and Tip                               //
//        .=.     .-=-.                                                       //
//       //"\\   ////"\\                                                      //
//      (/6 6\)  ( 6 6 )                                                      //
//      )\ = /(   \ - /                                                       //
//    _(_ ) ( _) _.) (._                                                      //
//   (_/ `\_/` \`  `:`  `\                                                    //
//    / (_ @ _) \|  :  |\ \                                                   //
//    \ \)___(/ /|  :  |/ /                                                   //
//     \/`"""`\/ \_ : _/ /                                                    //
//      |     |   |===|_)                                                     //
//      |     |   | L |                                                       //
//      |_____|   | | |                                                       //
//        |||     | | |                                                       //
//        |||     | | |                                                       //
//        |||     | | |                                                       //
//        |||     |_|_|                                                       //
//       / Y \    / T \                                                       //
//      `"`"`    `"`"`                                                        //
//                                                                            //
//         History update and review.                                         //
//         Ver. 1.0.0  2011-08-29  Lilly                                      //
//                                                                            //
////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////
unit MdOption;

interface
uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, ComCtrls, StdCtrls, Buttons, DB, ADODB, Registry,Excel2000,
  IdHashMessageDigest ,EncdDecd ,ExtCtrls ,StrUtils ,ComObj ;



  type

    TStringArray = array of string;

    ///////////// Procedure And Function
    ///
    function ConsiderName(ssName : string;SurName : string) : string ;
    procedure GetPic(var Im : TImage;Emp : string) ;
    procedure DecodeBaseToFile(const FileName: string;const EncodedString: AnsiString);
    function MD5_new(const src : string) : string;
    function FloatProtect(Edt : TEdit) : boolean ;
    function Xls_To_StringGrid(AGrid: TStringGrid; AXLSFile: string): Boolean;
    procedure ModifySectionName() ;
    function ConsiderDate(sDate : string) : string ;
    function ConsiderEmp(EmpNo : string) : string ;
    procedure ClearSgSectionList() ;
    function EncodeFile(const FileName: string): AnsiString;  stdcall;
    procedure UpdateVer() ;

implementation
uses aUsedlll , DmComponent ,MdGlobal ,FmMain ,FmLogin ,FmTemp ,MdMain  ,MdOptionNetwork ;

////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////                      //
//        Consider Name.(ตกแต่งชื่อนามสกุลให้มีตัวใหญง) //                      //
//                                                    //                      //
//        Ver. 1.0.0   2012-05-28  Lilly              //                      //
//                                                    //                      //
////////////////////////////////////////////////////////                      //
////////////////////////////////////////////////////////////////////////////////
function ConsiderName(ssName : string;SurName : string) : string ;
var s     : string  ;
    i,j     : integer ;
    sName   : string ;
    sSurname: string ;
begin
    /// ตกแต่งชื่อให้สวย
    s := '' ;
    for i := 2 to length(ssName) do
    begin
        s := s+ssName[i] ;
    end;
    sName := UpperCase(ssName[1])+LowerCase(s) ;

    /// ตกแต่งนามสกุลให้สวย
    s := '' ;
    for i := 2 to length(SurName) do
    begin
        s := s+SurName[i] ;
    end;
    sSurName := UpperCase(SurName[1])+LowerCase(s) ;

    Result := sName + ' ' + sSurName ;
end;

////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////                      //
//                                                    //                      //
//        Get Picture.                                //                      //
//                                                    //                      //
////////////////////////////////////////////////////////                      //
////////////////////////////////////////////////////////////////////////////////
procedure GetPic(var Im : TImage;Emp : string) ;
var DirRoot : string ;
    sDir  : string ;
    sPic  : string ;
    PicCall : string ;
    pic : pchar;
    i : integer ;
    Re : string ;
    Sql : string ;
begin
    GetDir(0,DirRoot) ;
    {$IOChecks off}
    MkDir('Config\PIC') ;
    {$IOChecks on}

    sPic    := Emp+'.jpg' ;
    sDir    := ExtractFilePath(DirRoot+'\Config\')+'PIC\' ;
    PicCall := sDir+sPic ;

    Sql := format('Select pic '+
                  'from picture_list '+
                  'where user_index = (select user_index '+
                                      'from user_list '+
                                      'where upper(id) = %s)',[QuotedStr(UpperCase(Emp))]) ;
    QueryOpen(Sql,DamComponent.ZQueryUser) ;

    Pic := PChar(DamComponent.ZQueryUser.FieldByName('pic').AsString)  ;
    //Pic := '' ;
    try
        DecodeBaseToFile(PicCall,pic) ;
        Im.Picture.LoadFromFile(PicCall) ;
    except
        Im.Picture.LoadFromFile(sDir+'Nidec.jpg') ;
    end;
end;

////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////                      //
//                                                    //                      //
//        Decode base picture to file.                //                      //
//                                                    //                      //
////////////////////////////////////////////////////////                      //
////////////////////////////////////////////////////////////////////////////////
procedure DecodeBaseToFile(const FileName: string;
const EncodedString: AnsiString);
var
bytes: TBytes;
Stream: TFileStream;
begin
    bytes := DecodeBase64(EncodedString);
    Stream := TFileStream.Create(FileName, fmCreate);
    try
    if bytes<>nil then
        Stream.WriteBuffer(bytes[0], Length(bytes));
    finally
        Stream.Free;
    end;
end;

////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////                      //
//                                                    //                      //
//        Md5 Encoder.                                //                      //
//                                                    //                      //
////////////////////////////////////////////////////////                      //
////////////////////////////////////////////////////////////////////////////////
function MD5_new(const src : string) : string;
var
   idmd5 : TIdHashMessageDigest5;
begin
   idmd5 := TIdHashMessageDigest5.Create;
   try
     result := idmd5.HashStringAsHex(UTF8ToString(src)) ;
   finally
     idmd5.Free;
   end;
end;

////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////                      //
//                                                    //                      //
//        Float input protection.                     //                      //
//                                                    //                      //
////////////////////////////////////////////////////////                      //
////////////////////////////////////////////////////////////////////////////////
function FloatProtect(Edt : TEdit) : boolean ;
var s : string ;
    lastt : string ;
    TempS : string ;
    PosCha : integer ;
    i,k : integer ;
    Err : boolean ;
    DotPass : boolean ;
begin
    TempS := '' ;
    DotPass := false ;
    s := Edt.Text ;
    err := false ;

    for i := 1 to length(s) do
    begin
        try
            StrToFloat(s[i]) ;
            if (s[i] = '.') then
            begin
                if (Dotpass = false) then
                begin
                    DotPass := true ;
                    TempS := TempS+s[i] ;
                end else
                begin
                    Err := true ;
                    k := i-1 ;
                end;
            end else
                TempS := TempS+s[i] ;
        except
              Err := true ;
              k := i-1 ;
        end;
    end;

    if Err = true then
    begin
        Edt.Text := Temps ;
        Edt.SelStart := k ;
        Result := false ;
    end else
        Result := true ;
end;

////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////
//                                                                            //
//        Read Excel file to sttringgrid.                                     //
//        Ver. 2020-06-22 Lilly.                                              //                                                                    //
//                                                                            //
////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////
function Xls_To_StringGrid(AGrid: TStringGrid; AXLSFile: string): Boolean;
const
    xlCellTypeLastCell = $0000000B;
var XLApp, Sheet: OLEVariant;
    RangeMatrix: Variant;
    x, y, k, r,i,j,a : Integer;

begin
    Cb[12].Items.Clear ;
    Cb[13].Items.Clear ;
    Cb[14].Items.Clear ;
    Cb[16].Items.Clear ;

    P[12].Color := $00FF8000 ;
    P[13].Color := $00FF8000 ;
    P[14].Color := $00FF8000 ;
    P[16].Color := $00FF8000 ;
    FrmMain.SgSectionList.RowCount := 2 ;

    Result := False;
    // Create Excel-OLE Object
    XLApp := CreateOleObject('Excel.Application');
    try
        // Hide Excel
        XLApp.Visible := False;
        // Open the Workbook
        XLApp.Workbooks.Open(AXLSFile);
        // Sheet := XLApp.Workbooks[1].WorkSheets[1];
        Sheet := XLApp.Workbooks[ExtractFileName(AXLSFile)].WorkSheets[1];
        // In order to know the dimension of the WorkSheet, i.e the number of rows
        // and the number of columns, we activate the last non-empty cell of it
        Sheet.Cells.SpecialCells(xlCellTypeLastCell, EmptyParam).Activate;
        // Get the value of the last row
        x := XLApp.ActiveCell.Row;
        // Get the value of the last column
        y := XLApp.ActiveCell.Column + 2+1 ;
        // Set Stringgrid's row &col dimensions.
        AGrid.RowCount := x;
        AGrid.ColCount := y;

        FrmTemp.SgList.RowCount := x;
        FrmTemp.SgList.ColCount := y;

        // Assign the Variant associated with the WorkSheet to the Delphi Variant
        RangeMatrix := XLApp.Range['A1', XLApp.Cells.Item[X, Y]].Value;
        //  Define the loop for filling in the TStringGrid
        k := 1;
    repeat
        for r := 1 to y do
        begin
            AGrid.Cells[(r+1), (k - 1)] := RangeMatrix[K, R];
//            AGrid.Cells[1,k] := IntToStr(k) ;

            FrmTemp.SgList.Cells[(r+1), (k - 1)] := RangeMatrix[K, R];
//            FrmTemp.SgList.Cells[1,k] := IntToStr(k) ;

            if r = 1 then
            begin
                AGrid.Cells[1,k] := IntToStr(k) ;
                FrmTemp.SgList.Cells[1,k] := IntToStr(k) ;

                if Length(RangeMatrix[K, R]) < 5 then
                begin
                    AGrid.Cells[(r+1), (k - 1)]          := ConsiderEmp(RangeMatrix[K, R]);
                    FrmTemp.SgList.Cells[(r+1), (k - 1)] := AGrid.Cells[(r+1), (k - 1)] ;
                end;
            end;

            ////////////////////////////////////////////////////////////////////////
            if (r = 13)or(r = 14)or(r = 16) then
            begin
                if k = 1 then
                begin
                    Cb[r].Items.Add('_All') ;
                end else
                begin
                    j := 0 ;
                    for i := 0 to Cb[r].Items.Count - 1 do
                    begin
                        if UpperCase(Cb[r].Items.Strings[i]) = UpperCase(AGrid.Cells[r,k-1]) then
                            Inc(j) ;
                    end;
                    if j = 0 then
                        Cb[r].Items.Add(AGrid.Cells[r,k-1]) ;
                end;
            end;
            if (r = 16){or(r = 13)} then
            begin

                if k = 1 then
                begin
                    Cb[12].Items.Add('_All') ;
                end else
                begin
                    j := 0 ;
                    for i := 0 to Cb[12].Items.Count - 1 do
                    begin
                        if UpperCase(Cb[12].Items.Strings[i]) = UpperCase(AGrid.Cells[12,k-1]) then
                            Inc(j) ;
                    end;

                    if j = 0 then
                    begin
                        Cb[12].Items.Add(AGrid.Cells[12,k-1]) ;
//                        Cb[13].Items.Add(AGrid.Cells[13,k-1]) ;

                        a := Cb[12].Items.Count ;
                        if a > 2 then
                        FrmMain.SgSectionList.RowCount := a ;

                        FrmMain.SgSectionList.Cells[1,a-1] := AGrid.Cells[12,k-1] ;
                        FrmMain.SgSectionList.Cells[2,a-1] := AGrid.Cells[13,k-1] ;
                        FrmMain.SgSectionList.Cells[3,a-1] := AGrid.Cells[16,k-1] ;
                    end;
                end;
            end;

            if r = 17 then
            begin
                AGrid.Cells[17,k-1] := GetPositionLevelIndex(AGrid.Cells[14,k-1]) ;
            end;
        end;

        ////////////////////////////////////////////////////////////////////////
        if k = 1 then
        begin
            AGrid.Cells[1,0] := 'No.' ;
            FrmTemp.SgList.Cells[1,0] := 'No.' ;

            P[12].Caption := AGrid.Cells[12,0] ;
            P[13].Caption := AGrid.Cells[13,0] ;
            P[14].Caption := AGrid.Cells[14,0] ;
            P[16].Caption := AGrid.Cells[16,0] ;
        end;


        Inc(k, 1);
        AGrid.RowCount := k + 1;
        FrmTemp.SgList.RowCount := k + 1;
        //Delay(0) ;
        if k mod 200 = 0 then
        begin
            AGrid.Row := k ;
            FrmTemp.SgList.Row := k ;
            Delay(0) ;
        end;
    until k > x;
    // Unassign the Delphi Variant Matrix
    RangeMatrix := Unassigned;

    AGrid.Row := k-1 ;
    FrmTemp.SgList.Row := k-1 ;
    finally
        // Quit Excel
        if not VarIsEmpty(XLApp) then
        begin
            // XLApp.DisplayAlerts := False;
          XLApp.Quit;
          XLAPP := Unassigned;
          Sheet := Unassigned;

          if AGrid.RowCount >= 4 then
          begin
              AGrid.RowCount := AGrid.RowCount - 2 ;
              FrmTemp.SgList.RowCount := FrmTemp.SgList.RowCount - 2 ;
          end;
          Result := True;
        end;
    end;
    FrmMain.SgList.Row := 1 ;
    ModifySectionName() ;
    Cb[12].Text := '_All' ;
    Cb[13].Text := '_All' ;
    Cb[14].Text := '_All' ;
    Cb[16].Text := '_All' ;
    AGrid.ColCount := 17 ;
end;

procedure ModifySectionName() ;
var i,j,k,l : integer ;
    AllCount ,FilterCount : integer ;
    s : string ;
    DupList : array of string ;
begin
    FrmMain.SgSectionList.Cells[4,0] := 'department_index' ;
    FrmMain.SgSectionList.Cells[5,0] := 'section_index' ;

    AllCount    := FrmMain.SgSectionList.RowCount - 1 ;
    FilterCount := FrmMain.CbSectionList.Items.Count - 1 ;
    if AllCount > FilterCount then
    begin
        l := 0 ;
        for i := 1 to FilterCount do
        begin
            k := 0 ;
            for j := 1 to AllCount do
            begin
                if FrmMain.CbSectionList.Items.Strings[i] = FrmMain.SgSectionList.Cells[2,j] then
                    Inc(k) ;

                if k > 1 then
                begin
                    Inc(l) ;
                    Setlength(DupList,l) ;
                    DupList[l-1] := FrmMain.CbSectionList.Items.Strings[i] ;
                    Break ;
                end;
            end;
        end;

        // Modify Dup Name.
        for i := 0 to l - 1 do
        begin
            for j := 1 to AllCount do
            begin
                if FrmMain.SgSectionList.Cells[2,j] = DupList[i] then
                begin
                    s := FrmMain.SgSectionList.Cells[2,j] ;
                    s := s + ' (' + FrmMain.SgSectionList.Cells[1,j]+')' ;
                    FrmMain.SgSectionList.Cells[2,j] := s ;
                end;
            end;
        end;
    end;
end;

function ConsiderDate(sDate : string) : string ;
var s,s1 : string  ;
begin
    s := sDate ;
    if (s[3] = '/')and(s[6] = '/') then
        s1 := s[7]+s[8]+s[9]+s[10]+'-'+s[4]+s[5]+'-'+s[1]+s[2]
    else
        s1 := s ;
    Result := s1 ;
end;


function ConsiderEmp(EmpNo : string) : string ;
var s,s1 : string ;
    i,j : integer ;
begin
    j := Length(EmpNo) ;
    if j < 5 then
    begin
        s := '' ;
        for i := 1 to 5-j do
            s := s+'0' ;

        Result := s+EmpNo
    end else
        Result := EmpNo ;
end;


procedure ClearSgSectionList() ;
var i,j : integer ;
begin
    for i := 1 to FrmMain.SgSectionList.ColCount-1 do
    begin
        for j := 1 to FrmMain.SgSectionList.RowCount-1 do
            FrmMain.SgSectionList.Cells[i,j] := '' ;
    end;
    FrmMain.SgSectionList.RowCount := 2 ;

    for i := 1 to FrmMain.SgPositionList.ColCount-1 do
    begin
        for j := 1 to FrmMain.SgPositionList.RowCount-1 do
            FrmMain.SgPositionList.Cells[i,j] := '' ;
    end;
    FrmMain.SgPositionList.RowCount := 2 ;
end;

function EncodeFile(const FileName: string): AnsiString;  stdcall;
var
    stream: TMemoryStream;
begin
    try
        stream := TMemoryStream.Create;
        try
            stream.LoadFromFile(Filename);
            result := EncodeBase64(stream.Memory, stream.Size);
        finally
            stream.Free;
        end;
    except

    end;
end;

////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////                      //
//                                                    //                      //
//        Update Version.                             //                      //
//                                                    //                      //
////////////////////////////////////////////////////////                      //
////////////////////////////////////////////////////////////////////////////////
procedure UpdateVer() ;
var MainDir,sever_dir,exename : string ;
    DirRoot : string ;
begin
    {DirRoot := ExtractFilePath(Application.ExeName) ;

    MainDir   := DirRoot ;
    sever_dir := 'K:\4.Other\0_TraceSystem\PanDa\PanDa2012.exe';
    exename   := 'PanDa2012.exe' ;

    checkupdate(application,MainDir,sever_dir,exename,1) ;}
end;
end.
