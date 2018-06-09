unit CHIStr;
(*******************************************************************************************
  Author      :Kees Hiemstra
  Version     :1.13 (July 5, 2003)
              - On RightStr(Line, 0), the wrong result was given.
              - Optimized memory use.
              - Added PosAt.
              - Added PosNum.
  History     :1.12 (March 9, 2003)
              - On LeftStr(Line, 0), the wrong result was given.
              :1.11 (February 26, 2003)
              - Added SplitSeparator.
              :1.10 (March 27, 2002)
              - Every function handles AnsiString.
              - Added BackPos.
              :1.04 (November 11, 2001)
              - Added EndStr.
              :1.03 (August 28, 2000)
              - Added IntToStrF.
              :1.02 (January 2, 1996)
              - Added LongInt2Base.
              :1.01 (November 6, 1995)
              - Added ReplaceStr.
              :1.00 (April 5, 1995)
              - Initial version.
*******************************************************************************************)

{$DEBUGINFO OFF}

Interface
Type
   Base36Str            =String[7];

Function BackPos(Const Search, Line:String):Integer;
Function LeftStr(Const Line:String; Start:Integer):String;
Function RightStr(Const Line:String; Start:Integer):String;
Function EndStr(Const Line:String; Start:Integer):String;
Function LTrim(Const Line:String):String;
Function RTrim(Const Line:String):String;
Function LongInt2Base36(Number:LongInt):Base36Str;
Function ReplaceStr(Const Line, ToReplace, ReplaceWith:String):String;
Function IntToStrF(i:LongInt):String;
Function SplitSeparator(Var Line:String; Var Split:String; Const Separator:String):Boolean;
Function PosAt(Const Search, Line:String; Start:Integer):Integer;
Function PosNum(Const Search, Line:String):Integer;

Implementation
Uses
  SysUtils;

Const
	Base36					:Array [0..35] Of Char
									= '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ';

(*#01***************************************************************************************
	BackPos(String1, String2) => Integer

    Returns the start position of String1 in String2 searching backwards

  Examples:

    BackPos('AABCAA',  'AA') => 5
    BackPos('AABCAA',  'AC') => 0
*******************************************************************************************)
Function BackPos(Const Search, Line:String):Integer;
Begin
  Result := Length(Line);
  If Result = 0 Then
    Exit;

  If Length(Search) = 1 Then
  Begin
    While (Result > 0) And (Line[Result] <> Search[1]) Do
      Dec(Result);
  End //If Char search
  Else
  Begin
    Result := Result - Length(Search);
    While (Result > 0) And (Copy(Line, Result, Length(Search)) = Search) Do
      Dec(Result);
  End; //Else String search
End;

(*#01***************************************************************************************
	LeftStr(String1, Integer1) => String

    Returns the left of String1 at the start given by Integer1

  Examples:

    LeftStr('ABCDEF',  0) => ''
    LeftStr('ABCDEF',  2) => 'AB'
    LeftStr('ABCDEF', -2) => 'ABCD'
*******************************************************************************************)
Function LeftStr(Const Line:String; Start:Integer):String;
Begin
  //Something to return?
	If Start = 0 Then
  Begin
   	Result := '';
   	Exit;
  End; //If Start

  //Start is counting from the left
	If Start > 0 Then
  Begin
		Result := Copy(Line, 1, Start);
    Exit;
  End; //If Start

  //Start is counting from the right
  Result := Copy(Line, 1, Length(Line) + Start);
End;

(*#01***************************************************************************************
	RightStr(String1, Integer1) => String

    Returns the right of String1 at the start given by Integer1

  Examples:

    RightStr('ABCDEF',  0) => ''
    RightStr('ABCDEF',  2) => 'EF'
    RightStr('ABCDEF', -2) => 'CDEF'
*******************************************************************************************)
Function RightStr(Const Line:String; Start:Integer):String;
Begin
  //Return all?
	If Start = 0 Then
  Begin
   	Result := '';
   	Exit;
  End; //If Start

  //Start is counting from the right
	If Start > 0 Then
  Begin
		Result := Copy(Line, Length(Line) - Start + 1, Start);
    Exit;
  End; //If Start

  //Start is counting from the left
  Result := Copy(Line, -Start, Length(Line) + Start + 1);
End;

(*#01***************************************************************************************
	EndStr(String1, Integer1) => String

    Returns the remaining of String1 from start given by Integer1

  Examples:

    EndStr('ABCDEF',  0) => 'ABCDEF'
    EndStr('ABCDEF',  2) => 'BCDEF'
*******************************************************************************************)
Function EndStr(Const Line:String; Start:Integer):String;
Begin
  //Return the complete line
	If Start = 0 Then
  Begin
   	EndStr := Line;
   	Exit;
  End; //If Start

  Result := Copy(Line, Start, Length(Line) - Start + 1);
End;

(*#01***************************************************************************************
  LTrim(String1) => String

    Returns String1 without leading spaces and tabs
*******************************************************************************************)
Function LTrim(Const Line:String):String;
Var
	Counter					      :Integer;
  Len                   :Integer;
Begin
	Counter := 1;
  Len := Length(Line);

	While (Counter <= Len) And (Line[Counter] In [#9, #32]) Do
		Inc(Counter);

	Result := Copy(Line, Counter, Len - Counter + 1);
End;

(*#01***************************************************************************************
	RTrim(String1) => String

    Returns String1 without trailing spaces and tabs
*******************************************************************************************)
Function RTrim(Const Line:String):String;
Var
	Counter					:Integer;
Begin
	Counter := Length(Line);

	While (Counter > 0) And (Line[Counter] In [#9, #32]) Do
		Dec(Counter);

	Result := Copy(Line, 1, Counter);
End;

(*#01***************************************************************************************
	LongInt2Base36(Integer1) => Base36Str

    Translates Integer1 to a number based on 36 characters (instead of 2 {binary},
    8 {octal}, 10 {decimal} or 16 {hexadecimal})

    Binairy system => base2
    Trinary system => base3
    Octaldecimal => base8
    Decimal => base10
    Hexadecimal => base16

    Places	Max
      1     36
      2     1.296
      3		  46.655
      4		  1.679.615
      5		  60.466.175
      6		  2.176.782.335
      7		  78.364.164.095

    Byte		255
    Word		65535
    LongInt	4.294.967.295
*******************************************************************************************)
Function LongInt2Base36(Number:LongInt):Base36Str;
Begin
	Result := '';

  While Number > 0 Do
  Begin
   	Result := Base36[Number Mod 36] + Result;
   	Number := Number Div 36;
  End; //While Number
End;

(*#01***************************************************************************************
 	ReplaceStr(String1, String2, String3) => String

    Replace in String1 all String2 with String3

  Example:

 		ReplaseStr('ABCDEFF', 'DEF', 'Q') => 'ABCQF'
*******************************************************************************************)
Function ReplaceStr(Const Line, ToReplace, ReplaceWith:String):String;
Var
  Position					    :Integer;
  Len						        :Integer;
Begin
 	Result := Line;

  //Is there anything to do
  If ToReplace = '' Then
  Begin
   	Exit;
  End;

  //Replaces from right to left all occurences
  Len := Length(ToReplace);
	Position := Length(Result) - Len + 1;
  While Position > 0 Do
  Begin
   	If Copy(Result, Position, Len) = ToReplace Then
    Begin
     	Delete(Result, Position, Len);
      Insert(ReplaceWith, Result, Position);
    End; //If Copy
    Dec(Position);
  End; //While Position
End;

Function IntToStrF(i:LongInt):String;
(*#01***************************************************************************************
  IntToStrF(Integer1) => String

    Returns a string from Integer1 with digitgrouping
*******************************************************************************************)
Var
	Len										:Integer;
Begin
	Result := IntToStr(i);
  Len := Length(Result) - 2;
  While Len > 1 Do
  Begin
    Insert('.', Result, Len);
  	Len := Len - 3;
  End; //While Len
End;

(*#01***************************************************************************************
	SplitSeparator(String1, String2, String3) => Boolean

    Splits String1 in String1 and String2 on separator String3. If the
    operation succeded, the function retuns true

  Examples:

    SplitSeparator('ABC;DE', Split, ';') => String1 = 'DE'
                                            String2 = 'ABC'
                                            Result = True
    SplitSeparator('ABC;DE', Split, '.') => String1 = ''
                                            String2 = 'ABC;DE'
                                            Result = False
*******************************************************************************************)
Function SplitSeparator(Var Line:String; Var Split:String; Const Separator:String):Boolean;
Var
  Position              :Integer;
Begin
  Position := Pos(Separator, Line);
  If Position > 0 Then
  Begin
    //Separator found
    Split := LeftStr(Line, Position - 1);
    Line := EndStr(Line, Position + 1);
    Result := True;
  End //If Position
  Else
  Begin
    //Separator NOT found
    Split := Line;
    Line := '';
    Result := False;
  End; //Else Position
End;

Function PosAt(Const Search, Line:String; Start:Integer):Integer;
(*#01***************************************************************************************
  PosAt(String1, String2, Integer1) => Integer

    Returns the start position of String1 in String2 searching from Integer1
*******************************************************************************************)
Begin
  Result := Pos(Search, Copy(Line, Start, Length(Line) - Start + 1));

  If Result > 0 Then
    Result := Result + Start - 1;
End;

Function PosNum(Const Search, Line:String):Integer;
(*#01***************************************************************************************
  PosNum(String1, String2):Integer

    Returns the number of occurences of String1 in String2
*******************************************************************************************)
Var
  Start                 :Integer;
Begin
  Result := 0;

  Start := Pos(Search, Line);
  While Start > 0 Do
  Begin
    Inc(Result);
    Start := PosAt(Search, Line, Start + 1);
  End; //While Start
End;

end.
