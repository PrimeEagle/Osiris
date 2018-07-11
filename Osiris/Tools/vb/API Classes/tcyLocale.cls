VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLocale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Declare Function GetThreadLocale Lib "kernel32" () As Long

Enum EMonth
    asNoMonth = 0
    asJanuary = 1
    asFebruary = 2
    asMarch = 3
    asApril = 4
    asMay = 5
    asJune = 6
    asJuly = 7
    asAugust = 8
    asSeptember = 9
    asOctober = 10
    asNovember = 11
    asDecember = 12
End Enum

Enum EDOW
    asNoDOW = 0
    asSunday = 1
    asMonday = 2
    asTuesday = 3
    asWednesday = 4
    asThursday = 5
    asFriday = 6
    asSaturday = 7
End Enum

Enum ELocal_Abbreviated
    CountryShort = &H7        '  abbreviated country name
    LanguageShort = &H3       '  abbreviated language name
    DateShort = &H1F          '  short date format string
    Mon = &H31                '  abbreviated name for Monday
    Tue = &H32                '  abbreviated name for Tuesday
    Wed = &H33                '  abbreviated name for Wednesday
    Thu = &H34                '  abbreviated name for Thursday
    Fri = &H35                '  abbreviated name for Friday
    Sat = &H36                '  abbreviated name for Saturday
    Sun = &H37                '  abbreviated name for Sunday
    Jan = &H44                '  abbreviated name for January
    Oct = &H4D                '  abbreviated name for October
    Nov = &H4E                '  abbreviated name for November
    Dec = &H4F                '  abbreviated name for December
    Feb = &H45                '  abbreviated name for February
    Mar = &H46                '  abbreviated name for March
    Apr = &H47                '  abbreviated name for April
    MayShort = &H48                '  abbreviated name for May
    Jun = &H49                '  abbreviated name for June
    Jul = &H4A                '  abbreviated name for July
    Aug = &H4B                '  abbreviated name for August
    Sep = &H4C                '  abbreviated name for September
End Enum

Enum ELocal
    CountryLong = &H6                   '  localized name of country
    CntryInEnglish = &H1002       '  English name of country
    LanguageLong = &H2                   '  localized name of language
    LangInEnglish = &H1001      '  English name of language
    CurrencySmbl = &H14           '  local monetary symbol
    DateSeparator = &H1D            '  date separator
    Monday = &H2A                   '  long name for Monday
    Tuesday = &H2B                  '  long name for Tuesday
    Wednesday = &H2C                '  long name for Wednesday
    Thursday = &H2D                 '  long name for Thursday
    Friday = &H2E                   '  long name for Friday
    Saturday = &H2F                 '  long name for Saturday
    Sunday = &H30                   '  long name for Sunday
    DateLong = &H20                 '  long date format string
    TimeFrmt = &H1003                '  time format string
    TimeSeparator = &H1E            '  time separator
    September = &H40                '  long name for September
    August = &H3F                   '  long name for August
    July = &H3E                     '  long name for July
    June = &H3D                     '  long name for June
    MayLong = &H3C                      '  long name for May
    April = &H3B                    '  long name for April
    March = &H3A                    '  long name for March
    February = &H39                 '  long name for February
    December = &H43                 '  long name for December
    November = &H42                 '  long name for November
    October = &H41                  '  long name for October
    January = &H38                  '  long name for January
End Enum

'Default Property Values:
Const m_def_Country = ""
Const m_def_DateFormat = ""
Const m_def_TimeFormat = ""
Const m_def_Month = ""
Const m_def_DayOfWeek = ""
Const m_def_Abbreviate = False
Const m_def_CurrencySymbol = ""
Const m_def_LanguageInEnglish = ""
Const m_def_CountryInEnglish = ""
Const m_def_Language = ""
Const m_def_FindMonth = 0
Const m_def_FindDayOfWeek = 0

'Property Variables:
Dim m_Country As String
Dim m_DateFormat As String
Dim m_TimeFormat As String
Dim m_Month As String
Dim m_DayOfWeek As String
Dim m_Abbreviate As Boolean
Dim m_CurrencySymbol As String
Dim m_LanguageInEnglish As String
Dim m_CountryInEnglish As String
Dim m_Language As String
Dim m_FindMonth As Long
Dim m_FindDayOfWeek As Long

Public Property Get Language() As String
    Language = m_Language
End Property

Private Sub Class_Initialize()
    m_CurrencySymbol = Space(99)
    m_Language = Space(99)
    m_Country = Space(99)
    m_DateFormat = Space(99)
    m_TimeFormat = Space(99)
    m_Month = Space(99)
    m_DayOfWeek = Space(99)
    m_Abbreviate = False
    m_LanguageInEnglish = Space(99)
    m_CountryInEnglish = Space(99)
    m_FindMonth = m_def_FindMonth
    m_FindDayOfWeek = m_def_FindDayOfWeek

    Refresh
End Sub
Public Property Get FindMonth() As EMonth
    FindMonth = m_FindMonth
End Property

Public Property Let FindMonth(ByVal New_FindMonth As EMonth)
    m_FindMonth = New_FindMonth
    
    Refresh
End Property

Public Property Get FindDayOfWeek() As EDOW
    FindDayOfWeek = m_FindDayOfWeek
End Property

Public Property Let FindDayOfWeek(ByVal New_FindDayOfWeek As EDOW)
    m_FindDayOfWeek = New_FindDayOfWeek
    
    Refresh
End Property

Private Function Locale_DOW_Long() As Long
    Dim Result As Long
    
    Select Case m_FindDayOfWeek
        Case Is = asSunday
            Result = Sunday
            
        Case Is = asMonday
            Result = Monday
            
        Case Is = asTuesday
            Result = Tuesday
            
        Case Is = asWednesday
            Result = Wednesday
        
        Case Is = asThursday
            Result = Thursday
            
        Case Is = asFriday
            Result = Friday
        
        Case Is = asSaturday
            Result = Saturday
    End Select
    
    Locale_DOW_Long = Result
End Function

Private Function Locale_DOW_Short() As Long
    Dim Result As Long
    
    Select Case m_FindDayOfWeek
        Case Is = asSunday
            Result = Sun
            
        Case Is = asMonday
            Result = Mon
            
        Case Is = asTuesday
            Result = Tue
            
        Case Is = asWednesday
            Result = Wed
        
        Case Is = asThursday
            Result = Thu
            
        Case Is = asFriday
            Result = Fri
        
        Case Is = asSaturday
            Result = Sat
    End Select
    
    Locale_DOW_Short = Result
End Function

Private Function Locale_Month_Long() As Long
    Dim Result As Long
    
    Select Case m_FindMonth
        Case Is = asJanuary
            Result = January
            
        Case Is = asFebruary
            Result = February
            
        Case Is = asMarch
            Result = March
            
        Case Is = asApril
            Result = April
        
        Case Is = asMay
            Result = MayLong
            
        Case Is = asJune
            Result = June
        
        Case Is = asJuly
            Result = July
        
        Case Is = asAugust
            Result = August
        
        Case Is = asSeptember
            Result = September
        
        Case Is = asOctober
            Result = October
        
        Case Is = asNovember
            Result = November
    
        Case Is = asDecember
            Result = December
    End Select
    
    Locale_Month_Long = Result
End Function

Private Function Locale_Month_Short() As Long
    Dim Result As Long
    
    Select Case m_FindMonth
        Case Is = asJanuary
            Result = Jan
            
        Case Is = asFebruary
            Result = Feb
            
        Case Is = asMarch
            Result = Mar
            
        Case Is = asApril
            Result = Apr
        
        Case Is = asMay
            Result = MayShort
            
        Case Is = asJune
            Result = Jun
        
        Case Is = asJuly
            Result = July
        
        Case Is = asAugust
            Result = Aug
        
        Case Is = asSeptember
            Result = Sep
        
        Case Is = asOctober
            Result = Oct
        
        Case Is = asNovember
            Result = Nov
    
        Case Is = asDecember
            Result = Dec
    End Select
    
    Locale_Month_Short = Result
End Function

Public Sub Refresh()
    Dim Result   As Long
    Dim LocaleID As Long
    
    LocaleID = GetThreadLocale
    
    Result = GetLocaleInfo(LocaleID, CurrencySmbl, m_CurrencySymbol, 99)
    Result = GetLocaleInfo(LocaleID, TimeFrmt, m_TimeFormat, 99)
    Result = GetLocaleInfo(LocaleID, LangInEnglish, m_LanguageInEnglish, 99)
    Result = GetLocaleInfo(LocaleID, CntryInEnglish, m_CountryInEnglish, 99)
    
    If m_Abbreviate Then
        Result = GetLocaleInfo(LocaleID, CountryShort, m_Country, 99)
        Result = GetLocaleInfo(LocaleID, Locale_DOW_Short(), m_DayOfWeek, 99)
        Result = GetLocaleInfo(LocaleID, Locale_Month_Short(), m_Month, 99)
        Result = GetLocaleInfo(LocaleID, LanguageShort, m_Language, 99)
        Result = GetLocaleInfo(LocaleID, DateShort, m_DateFormat, 99)
    Else
        Result = GetLocaleInfo(LocaleID, CountryLong, m_Country, 99)
        Result = GetLocaleInfo(LocaleID, Locale_DOW_Long(), m_DayOfWeek, 99)
        Result = GetLocaleInfo(LocaleID, Locale_Month_Long(), m_Month, 99)
        Result = GetLocaleInfo(LocaleID, LanguageLong, m_Language, 99)
        Result = GetLocaleInfo(LocaleID, DateLong, m_DateFormat, 99)
    End If
End Sub
Public Property Get Country() As String
    Country = m_Country
End Property

Public Property Get DateFormat() As String
    DateFormat = m_DateFormat
End Property

Public Property Get TimeFormat() As String
    TimeFormat = m_TimeFormat
End Property

Public Property Get Month() As String
    Month = m_Month
End Property

Public Property Get DayOfWeek() As String
    DayOfWeek = m_DayOfWeek
End Property

Public Property Get Abbreviate() As Boolean
    Abbreviate = m_Abbreviate
End Property

Public Property Let Abbreviate(ByVal New_Abbreviate As Boolean)
    m_Abbreviate = New_Abbreviate
    
    Refresh
End Property

Public Property Get CurrencySymbol() As String
    CurrencySymbol = m_CurrencySymbol
End Property

Public Property Get LanguageInEnglish() As String
    LanguageInEnglish = m_LanguageInEnglish
End Property

Public Property Get CountryInEnglish() As String
    CountryInEnglish = m_CountryInEnglish
End Property


