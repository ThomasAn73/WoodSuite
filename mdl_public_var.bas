Attribute VB_Name = "mdl_public_var"
'These two are the work-horses; array and collection
Public LumberSizes() As Single 'holds wood sizes information
Public tree As New matrix 'Holds wood species and grades information
Public Const PI As Double = 3.14159265358979
Type specs
    Area As Single
    FsubB As Single
    FsubC As Single
    FsubT As Single
    MaxMoment As Double
    MaxShear As Double
    MaxUnbracedLength As Single
    SectMod As Single
    ElastMod As Single
    beamWeight As Single
    beamLength As Single
    beamThickness As Single
    beamWidth As Single
    MaxShearLoc As Double
    MaxMomentLoc As Double
    UnbracedLengthLoad As Single
    compression As Boolean
End Type
Public currspecs As specs

Public Const delims As String = " X"
Public sizes_file As String
Public species_file As String

Public Const cutter As Integer = 182 'Delimiter for frm_editwood
Public Const TheFormat As String = "######0.0"
Public Enum err
    notnumber = 0
    Exists = 1
    nonselect = 2
    overflow = 3
    overchange = 4
    no_entry = 5
    empty_field = 6
    proper_size = 7
    exceed_length = 8
    noReactions = 9
    noEvaluate = 10
    listSupportFailure = 11
    noinclusion1 = 12
    noinclusion2 = 13
    zeroMag = 14
    noSpecies = 15
    locationNonUnique = 16
    noLoads = 17
    MinusSpan = 18
    Minusfield = 19
End Enum
Public Enum answer
    ok = 1
    yes = 6
    no = 7
    Cancel = 2
End Enum
Public Enum ask
    sure = 0
    redirect = 1
    sure2 = 2
    Update = 3
    Save = 4
    clear_all = 5
End Enum

' identifiers for the contents of the list box that contains "Loads" information
Public Enum L
    span = 2
    location = 3
    magnitude = 4
    direction = 5
End Enum
Type FactorsSelect
    flat As Boolean
    wet As Boolean
    duration As Boolean
End Type
Public factor As FactorsSelect

Type corrFactors
    f As Single
    f_Fc As Single
    r As Single
    fu As Single
    m As Single
    m_Fc As Single
    m_E As Single
    d As Single
End Type
Public c As corrFactors

Type RGB_Palete
    Red As Integer
    Green As Integer
    Blue As Integer
End Type
