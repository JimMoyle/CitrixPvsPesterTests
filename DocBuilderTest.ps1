#region setup
    #Word build script
    Get-Process -Name *word* | Stop-Process

    #try and fix the issue with the $CompanyName variable
$Script:CoName = 'JimMoyle Ltd'
$UserName = 'Jim Moyle'
$Script:Title = 'Jim Title Test'
$SubjectTitle = 'Jim Subject Title Test'

    Write-Verbose "$(Get-Date): CoName is $($Script:CoName)"

    $MSWORD = $true

    $Script:FileName1 = 'C:\JimM\Test.docx'

    $CoverPage='Sideline'

    #the following values were attained from
    #http://groovy.codehaus.org/modules/scriptom/1.6.0/scriptom-office-2K3-tlb/apidocs/
    #http://msdn.microsoft.com/en-us/library/office/aa211923(v=office.11).aspx
    [int]$wdAlignPageNumberRight = 2
    [long]$wdColorGray15 = 14277081
    [long]$wdColorGray05 = 15987699
    [int]$wdMove = 0
    [int]$wdSeekMainDocument = 0
    [int]$wdSeekPrimaryFooter = 4
    [int]$wdStory = 6
    [int]$wdColorRed = 255
    [int]$wdColorBlack = 0
    [int]$wdWord2007 = 12
    [int]$wdWord2010 = 14
    [int]$wdWord2013 = 15
    [int]$wdWord2016 = 16
    [int]$wdFormatDocumentDefault = 16
    [int]$wdFormatPDF = 17
    #http://blogs.technet.com/b/heyscriptingguy/archive/2006/03/01/how-can-i-right-align-a-single-column-in-a-word-table.aspx
    #http://msdn.microsoft.com/en-us/library/office/ff835817%28v=office.15%29.aspx
    [int]$wdAlignParagraphLeft = 0
    [int]$wdAlignParagraphCenter = 1
    [int]$wdAlignParagraphRight = 2
    #http://msdn.microsoft.com/en-us/library/office/ff193345%28v=office.15%29.aspx
    [int]$wdCellAlignVerticalTop = 0
    [int]$wdCellAlignVerticalCenter = 1
    [int]$wdCellAlignVerticalBottom = 2
    #http://msdn.microsoft.com/en-us/library/office/ff844856%28v=office.15%29.aspx
    [int]$wdAutoFitFixed = 0
    [int]$wdAutoFitContent = 1
    [int]$wdAutoFitWindow = 2
    #http://msdn.microsoft.com/en-us/library/office/ff821928%28v=office.15%29.aspx
    [int]$wdAdjustNone = 0
    [int]$wdAdjustProportional = 1
    [int]$wdAdjustFirstColumn = 2
    [int]$wdAdjustSameWidth = 3

    [int]$PointsPerTabStop = 36
    [int]$Indent0TabStops = 0 * $PointsPerTabStop
    [int]$Indent1TabStops = 1 * $PointsPerTabStop
    [int]$Indent2TabStops = 2 * $PointsPerTabStop
    [int]$Indent3TabStops = 3 * $PointsPerTabStop
    [int]$Indent4TabStops = 4 * $PointsPerTabStop

    # http://www.thedoctools.com/index.php?show=wt_style_names_english_danish_german_french
    [int]$wdStyleHeading1 = -2
    [int]$wdStyleHeading2 = -3
    [int]$wdStyleHeading3 = -4
    [int]$wdStyleHeading4 = -5
    [int]$wdStyleNoSpacing = -158
    [int]$wdTableGrid = -155

    #http://groovy.codehaus.org/modules/scriptom/1.6.0/scriptom-office-2K3-tlb/apidocs/org/codehaus/groovy/scriptom/tlb/office/word/WdLineStyle.html
    [int]$wdLineStyleNone = 0
    [int]$wdLineStyleSingle = 1

    [int]$wdHeadingFormatTrue = -1
    [int]$wdHeadingFormatFalse = 0
#endregion

Function SetWordHashTable {
    Param([string]$CultureCode)

    #optimized by Michael B. SMith

    # DE and FR translations for Word 2010 by Vladimir Radojevic
    # Vladimir.Radojevic@Commerzreal.com

    # DA translations for Word 2010 by Thomas Daugaard
    # Citrix Infrastructure Specialist at edgemo A/S

    # CA translations by Javier Sanchez
    # CEO & Founder 101 Consulting

    #ca - Catalan
    #da - Danish
    #de - German
    #en - English
    #es - Spanish
    #fi - Finnish
    #fr - French
    #nb - Norwegian
    #nl - Dutch
    #pt - Portuguese
    #sv - Swedish
    #zh - Chinese

    [string]$toc = $(
        Switch ($CultureCode) {
            'ca-'	{ 'Taula automática 2'; Break }
            'da-'	{ 'Automatisk tabel 2'; Break }
            'de-'	{ 'Automatische Tabelle 2'; Break }
            'en-'	{ 'Automatic Table 2'; Break }
            'es-'	{ 'Tabla automática 2'; Break }
            'fi-'	{ 'Automaattinen taulukko 2'; Break }
            'fr-'	{ 'Table automatique 2'; Break } #changed 13-feb-2017 david roquier and samuel legrand
            'nb-'	{ 'Automatisk tabell 2'; Break }
            'nl-'	{ 'Automatische inhoudsopgave 2'; Break }
            'pt-'	{ 'Sumário Automático 2'; Break }
            'sv-'	{ 'Automatisk innehållsförteckning2'; Break }
            'zh-'	{ '自动目录 2'; Break }
        }
    )

    $Script:myHash = @{}
    $Script:myHash.Word_TableOfContents = $toc
    $Script:myHash.Word_NoSpacing = $wdStyleNoSpacing
    $Script:myHash.Word_Heading1 = $wdStyleheading1
    $Script:myHash.Word_Heading2 = $wdStyleheading2
    $Script:myHash.Word_Heading3 = $wdStyleheading3
    $Script:myHash.Word_Heading4 = $wdStyleheading4
    $Script:myHash.Word_TableGrid = $wdTableGrid
}

Function GetCulture {
    Param([int]$WordValue)

    #codes obtained from http://support.microsoft.com/kb/221435
    #http://msdn.microsoft.com/en-us/library/bb213877(v=office.12).aspx
    $CatalanArray = 1027
    $ChineseArray = 2052, 3076, 5124, 4100
    $DanishArray = 1030
    $DutchArray = 2067, 1043
    $EnglishArray = 3081, 10249, 4105, 9225, 6153, 8201, 5129, 13321, 7177, 11273, 2057, 1033, 12297
    $FinnishArray = 1035
    $FrenchArray = 2060, 1036, 11276, 3084, 12300, 5132, 13324, 6156, 8204, 10252, 7180, 9228, 4108
    $GermanArray = 1031, 3079, 5127, 4103, 2055
    $NorwegianArray = 1044, 2068
    $PortugueseArray = 1046, 2070
    $SpanishArray = 1034, 11274, 16394, 13322, 9226, 5130, 7178, 12298, 17418, 4106, 18442, 19466, 6154, 15370, 10250, 20490, 3082, 14346, 8202
    $SwedishArray = 1053, 2077

    #ca - Catalan
    #da - Danish
    #de - German
    #en - English
    #es - Spanish
    #fi - Finnish
    #fr - French
    #nb - Norwegian
    #nl - Dutch
    #pt - Portuguese
    #sv - Swedish
    #zh - Chinese

    Switch ($WordValue) {
        {$CatalanArray -contains $_} {$CultureCode = "ca-"}
        {$ChineseArray -contains $_} {$CultureCode = "zh-"}
        {$DanishArray -contains $_} {$CultureCode = "da-"}
        {$DutchArray -contains $_} {$CultureCode = "nl-"}
        {$EnglishArray -contains $_} {$CultureCode = "en-"}
        {$FinnishArray -contains $_} {$CultureCode = "fi-"}
        {$FrenchArray -contains $_} {$CultureCode = "fr-"}
        {$GermanArray -contains $_} {$CultureCode = "de-"}
        {$NorwegianArray -contains $_} {$CultureCode = "nb-"}
        {$PortugueseArray -contains $_} {$CultureCode = "pt-"}
        {$SpanishArray -contains $_} {$CultureCode = "es-"}
        {$SwedishArray -contains $_} {$CultureCode = "sv-"}
        Default {$CultureCode = "en-"}
    }

    Return $CultureCode
}

Function ValidateCoverPage {
    Param([int]$xWordVersion, [string]$xCP, [string]$CultureCode)

    $xArray = ""

    Switch ($CultureCode) {
        'ca-'	{
            If ($xWordVersion -eq $wdWord2016) {
                $xArray = ("Austin", "En bandes", "Faceta", "Filigrana",
                    "Integral", "Ió (clar)", "Ió (fosc)", "Línia lateral",
                    "Moviment", "Quadrícula", "Retrospectiu", "Sector (clar)",
                    "Sector (fosc)", "Semàfor", "Visualització principal", "Whisp")
            }
            ElseIf ($xWordVersion -eq $wdWord2013) {
                $xArray = ("Austin", "En bandes", "Faceta", "Filigrana",
                    "Integral", "Ió (clar)", "Ió (fosc)", "Línia lateral",
                    "Moviment", "Quadrícula", "Retrospectiu", "Sector (clar)",
                    "Sector (fosc)", "Semàfor", "Visualització", "Whisp")
            }
            ElseIf ($xWordVersion -eq $wdWord2010) {
                $xArray = ("Alfabet", "Anual", "Austin", "Conservador",
                    "Contrast", "Cubicles", "Diplomàtic", "Exposició",
                    "Línia lateral", "Mod", "Mosiac", "Moviment", "Paper de diari",
                    "Perspectiva", "Piles", "Quadrícula", "Sobri",
                    "Transcendir", "Trencaclosques")
            }
        }

        'da-'	{
            If ($xWordVersion -eq $wdWord2016) {
                $xArray = ("Austin", "BevægElse", "Brusen", "Facet", "Filigran",
                    "Gitter", "Integral", "Ion (lys)", "Ion (mørk)",
                    "Retro", "Semafor", "Sidelinje", "Stribet",
                    "Udsnit (lys)", "Udsnit (mørk)", "Visningsmaster")
            }
            ElseIf ($xWordVersion -eq $wdWord2013) {
                $xArray = ("BevægElse", "Brusen", "Ion (lys)", "Filigran",
                    "Retro", "Semafor", "Visningsmaster", "Integral",
                    "Facet", "Gitter", "Stribet", "Sidelinje", "Udsnit (lys)",
                    "Udsnit (mørk)", "Ion (mørk)", "Austin")
            }
            ElseIf ($xWordVersion -eq $wdWord2010) {
                $xArray = ("BevægElse", "Moderat", "Perspektiv", "Firkanter",
                    "Overskrid", "Alfabet", "Kontrast", "Stakke", "Fliser", "Gåde",
                    "Gitter", "Austin", "Eksponering", "Sidelinje", "Enkel",
                    "Nålestribet", "Årlig", "Avispapir", "Tradionel")
            }
        }

        'de-'	{
            If ($xWordVersion -eq $wdWord2016) {
                $xArray = ("Austin", "Bewegung", "Facette", "Filigran",
                    "Gebändert", "Integral", "Ion (dunkel)", "Ion (hell)",
                    "Pfiff", "Randlinie", "Raster", "Rückblick",
                    "Segment (dunkel)", "Segment (hell)", "Semaphor",
                    "ViewMaster")
            }
            ElseIf ($xWordVersion -eq $wdWord2013) {
                $xArray = ("Semaphor", "Segment (hell)", "Ion (hell)",
                    "Raster", "Ion (dunkel)", "Filigran", "Rückblick", "Pfiff",
                    "ViewMaster", "Segment (dunkel)", "Verbunden", "Bewegung",
                    "Randlinie", "Austin", "Integral", "Facette")
            }
            ElseIf ($xWordVersion -eq $wdWord2010) {
                $xArray = ("Alphabet", "Austin", "Bewegung", "Durchscheinend",
                    "Herausgestellt", "Jährlich", "Kacheln", "Kontrast", "Kubistisch",
                    "Modern", "Nadelstreifen", "Perspektive", "Puzzle", "Randlinie",
                    "Raster", "Schlicht", "Stapel", "Traditionell", "Zeitungspapier")
            }
        }

        'en-'	{
            If ($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016) {
                $xArray = ("Austin", "Banded", "Facet", "Filigree", "Grid",
                    "Integral", "Ion (Dark)", "Ion (Light)", "Motion", "Retrospect",
                    "Semaphore", "Sideline", "Slice (Dark)", "Slice (Light)", "ViewMaster",
                    "Whisp")
            }
            ElseIf ($xWordVersion -eq $wdWord2010) {
                $xArray = ("Alphabet", "Annual", "Austere", "Austin", "Conservative",
                    "Contrast", "Cubicles", "Exposure", "Grid", "Mod", "Motion", "Newsprint",
                    "Perspective", "Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend")
            }
        }

        'es-'	{
            If ($xWordVersion -eq $wdWord2016) {
                $xArray = ("Austin", "Con bandas", "Cortar (oscuro)", "Cuadrícula",
                    "Whisp", "Faceta", "Filigrana", "Integral", "Ion (claro)",
                    "Ion (oscuro)", "Línea lateral", "Movimiento", "Retrospectiva",
                    "Semáforo", "Slice (luz)", "Vista principal", "Whisp")
            }
            ElseIf ($xWordVersion -eq $wdWord2013) {
                $xArray = ("Whisp", "Vista principal", "Filigrana", "Austin",
                    "Slice (luz)", "Faceta", "Semáforo", "Retrospectiva", "Cuadrícula",
                    "Movimiento", "Cortar (oscuro)", "Línea lateral", "Ion (oscuro)",
                    "Ion (claro)", "Integral", "Con bandas")
            }
            ElseIf ($xWordVersion -eq $wdWord2010) {
                $xArray = ("Alfabeto", "Anual", "Austero", "Austin", "Conservador",
                    "Contraste", "Cuadrícula", "Cubículos", "Exposición", "Línea lateral",
                    "Moderno", "Mosaicos", "Movimiento", "Papel periódico",
                    "Perspectiva", "Pilas", "Puzzle", "Rayas", "Sobrepasar")
            }
        }

        'fi-'	{
            If ($xWordVersion -eq $wdWord2016) {
                $xArray = ("Filigraani", "Integraali", "Ioni (tumma)",
                    "Ioni (vaalea)", "Opastin", "Pinta", "Retro", "Sektori (tumma)",
                    "Sektori (vaalea)", "Vaihtuvavärinen", "ViewMaster", "Austin",
                    "Kuiskaus", "Liike", "Ruudukko", "Sivussa")
            }
            ElseIf ($xWordVersion -eq $wdWord2013) {
                $xArray = ("Filigraani", "Integraali", "Ioni (tumma)",
                    "Ioni (vaalea)", "Opastin", "Pinta", "Retro", "Sektori (tumma)",
                    "Sektori (vaalea)", "Vaihtuvavärinen", "ViewMaster", "Austin",
                    "Kiehkura", "Liike", "Ruudukko", "Sivussa")
            }
            ElseIf ($xWordVersion -eq $wdWord2010) {
                $xArray = ("Aakkoset", "Askeettinen", "Austin", "Kontrasti",
                    "Laatikot", "Liike", "Liituraita", "Mod", "Osittain peitossa",
                    "Palapeli", "Perinteinen", "Perspektiivi", "Pinot", "Ruudukko",
                    "Ruudut", "Sanomalehtipaperi", "Sivussa", "Vuotuinen", "Ylitys")
            }
        }

        'fr-'	{
            If ($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016) {
                $xArray = ("À bandes", "Austin", "Facette", "Filigrane",
                    "Guide", "Intégrale", "Ion (clair)", "Ion (foncé)",
                    "Lignes latérales", "Quadrillage", "Rétrospective", "Secteur (clair)",
                    "Secteur (foncé)", "Sémaphore", "ViewMaster", "Whisp")
            }
            ElseIf ($xWordVersion -eq $wdWord2010) {
                $xArray = ("Alphabet", "Annuel", "Austère", "Austin",
                    "Blocs empilés", "Classique", "Contraste", "Emplacements de bureau",
                    "Exposition", "Guide", "Ligne latérale", "Moderne",
                    "Mosaïques", "Mots croisés", "Papier journal", "Perspective",
                    "Quadrillage", "Rayures fines", "Transcendant")
            }
        }

        'nb-'	{
            If ($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016) {
                $xArray = ("Austin", "BevegElse", "Dempet", "Fasett", "Filigran",
                    "Integral", "Ion (lys)", "Ion (mørk)", "Retrospekt", "Rutenett",
                    "Sektor (lys)", "Sektor (mørk)", "Semafor", "Sidelinje", "Stripet",
                    "ViewMaster")
            }
            ElseIf ($xWordVersion -eq $wdWord2010) {
                $xArray = ("Alfabet", "Årlig", "Avistrykk", "Austin", "Avlukker",
                    "BevegElse", "Engasjement", "Enkel", "Fliser", "Konservativ",
                    "Kontrast", "Mod", "Perspektiv", "Puslespill", "Rutenett", "Sidelinje",
                    "Smale striper", "Stabler", "Transcenderende")
            }
        }

        'nl-'	{
            If ($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016) {
                $xArray = ("Austin", "Beweging", "Facet", "Filigraan", "Gestreept",
                    "Integraal", "Ion (donker)", "Ion (licht)", "Raster",
                    "Segment (Light)", "Semafoor", "Slice (donker)", "Spriet",
                    "Terugblik", "Terzijde", "ViewMaster")
            }
            ElseIf ($xWordVersion -eq $wdWord2010) {
                $xArray = ("Aantrekkelijk", "Alfabet", "Austin", "Bescheiden",
                    "Beweging", "Blikvanger", "Contrast", "Eenvoudig", "Jaarlijks",
                    "Krantenpapier", "Krijtstreep", "Kubussen", "Mod", "Perspectief",
                    "Puzzel", "Raster", "Stapels",
                    "Tegels", "Terzijde")
            }
        }

        'pt-'	{
            If ($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016) {
                $xArray = ("Animação", "Austin", "Em Tiras", "Exibição Mestra",
                    "Faceta", "Fatia (Clara)", "Fatia (Escura)", "Filete", "Filigrana",
                    "Grade", "Integral", "Íon (Claro)", "Íon (Escuro)", "Linha Lateral",
                    "Retrospectiva", "Semáforo")
            }
            ElseIf ($xWordVersion -eq $wdWord2010) {
                $xArray = ("Alfabeto", "Animação", "Anual", "Austero", "Austin", "Baias",
                    "Conservador", "Contraste", "Exposição", "Grade", "Ladrilhos",
                    "Linha Lateral", "Listras", "Mod", "Papel Jornal", "Perspectiva", "Pilhas",
                    "Quebra-cabeça", "Transcend")
            }
        }

        'sv-'	{
            If ($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016) {
                $xArray = ("Austin", "Band", "Fasett", "Filigran", "Integrerad", "Jon (ljust)",
                    "Jon (mörkt)", "Knippe", "Rutnät", "RörElse", "Sektor (ljus)", "Sektor (mörk)",
                    "Semafor", "Sidlinje", "VisaHuvudsida", "Återblick")
            }
            ElseIf ($xWordVersion -eq $wdWord2010) {
                $xArray = ("Alfabetmönster", "Austin", "Enkelt", "Exponering", "Konservativt",
                    "Kontrast", "Kritstreck", "Kuber", "Perspektiv", "Plattor", "Pussel", "Rutnät",
                    "RörElse", "Sidlinje", "Sobert", "Staplat", "Tidningspapper", "Årligt",
                    "Övergående")
            }
        }

        <# 'zh-'	{
            If ($xWordVersion -eq $wdWord2010 -or $xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016) {
                $xArray = ('奥斯汀', '边线型', '花丝', '怀旧', '积分',
                    '离子(浅色)', '离子(深色)', '母版型', '平面', '切片(浅色)',
                    '切片(深色)', '丝状', '网格', '镶边', '信号灯',
                    '运动型')
            }
        }
        #>

        Default	{
            If ($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016) {
                $xArray = ("Austin", "Banded", "Facet", "Filigree", "Grid",
                    "Integral", "Ion (Dark)", "Ion (Light)", "Motion", "Retrospect",
                    "Semaphore", "Sideline", "Slice (Dark)", "Slice (Light)", "ViewMaster",
                    "Whisp")
            }
            ElseIf ($xWordVersion -eq $wdWord2010) {
                $xArray = ("Alphabet", "Annual", "Austere", "Austin", "Conservative",
                    "Contrast", "Cubicles", "Exposure", "Grid", "Mod", "Motion", "Newsprint",
                    "Perspective", "Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend")
            }
        }
    }

    If ($xArray -contains $xCP) {
        $xArray = $Null
        Return $True
    }
    Else {
        $xArray = $Null
        Return $False
    }
}

Function CheckWordPrereq {
    If ((Test-Path  REGISTRY::HKEY_CLASSES_ROOT\Word.Application) -eq $False) {
        $ErrorActionPreference = $SaveEAPreference
        Write-Host "`n`n`t`tThis script directly outputs to Microsoft Word, please install Microsoft Word`n`n"
        Exit
    }

    #find out our session (usually "1" except on TS/RDC or Citrix)
    $SessionID = (Get-Process -PID $PID).SessionId

    #Find out if winword is running in our session
    [bool]$wordrunning = ((Get-Process 'WinWord' -ea 0)|? {$_.SessionId -eq $SessionID}) -ne $Null
    If ($wordrunning) {
        $ErrorActionPreference = $SaveEAPreference
        Write-Host "`n`n`tPlease close all instances of Microsoft Word before running this report.`n`n"
        Exit
    }
}

Function ValidateCompanyName {
    [bool]$xResult = Test-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "CompanyName"
    If ($xResult) {
        Return Get-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "CompanyName"
    }
    Else {
        $xResult = Test-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "Company"
        If ($xResult) {
            Return Get-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "Company"
        }
        Else {
            Return ""
        }
    }
}

Function _SetDocumentProperty {
    #jeff hicks
    Param([object]$Properties, [string]$Name, [string]$Value)
    #get the property object
    $prop = $properties | ForEach-Object {
        $propname = $_.GetType().InvokeMember("Name", "GetProperty", $Null, $_, $Null)
        If ($propname -eq $Name) {
            Return $_
        }
    } #ForEach

    #set the value
    $Prop.GetType().InvokeMember("Value", "SetProperty", $Null, $prop, $Value)
}

Function FindWordDocumentEnd {
    #return focus to main document
    $Script:Doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument
    #move to the end of the current document
    $Script:Selection.EndKey($wdStory, $wdMove) | Out-Null
}

Function SetupWord {
    Write-Verbose "$(Get-Date): Setting up Word"

    # Setup word for output
    Write-Verbose "$(Get-Date): Create Word comObject."
    $Script:Word = New-Object -comobject "Word.Application" -EA 0 4>$Null

    If (!$? -or $Null -eq $Script:Word) {
        Write-Warning "The Word object could not be created.  You may need to repair your Word installation."
        $ErrorActionPreference = $SaveEAPreference
        Write-Error "`n`n`t`tThe Word object could not be created.  You may need to repair your Word installation.`n`n`t`tScript cannot continue.`n`n"
        Exit
    }

    Write-Verbose "$(Get-Date): Determine Word language value"
    If ( ( validStateProp $Script:Word Language Value__ ) ) {
        [int]$Script:WordLanguageValue = [int]$Script:Word.Language.Value__
    }
    Else {
        [int]$Script:WordLanguageValue = [int]$Script:Word.Language
    }

    If (!($Script:WordLanguageValue -gt -1)) {
        $ErrorActionPreference = $SaveEAPreference
        Write-Error "`n`n`t`tUnable to determine the Word language value.`n`n`t`tScript cannot continue.`n`n"
        AbortScript
    }
    Write-Verbose "$(Get-Date): Word language value is $($Script:WordLanguageValue)"

    $Script:WordCultureCode = GetCulture $Script:WordLanguageValue

    SetWordHashTable $Script:WordCultureCode

    [int]$Script:WordVersion = [int]$Script:Word.Version
    If ($Script:WordVersion -eq $wdWord2016) {
        $Script:WordProduct = "Word 2016"
    }
    ElseIf ($Script:WordVersion -eq $wdWord2013) {
        $Script:WordProduct = "Word 2013"
    }
    ElseIf ($Script:WordVersion -eq $wdWord2010) {
        $Script:WordProduct = "Word 2010"
    }
    ElseIf ($Script:WordVersion -eq $wdWord2007) {
        $ErrorActionPreference = $SaveEAPreference
        Write-Error "`n`n`t`tMicrosoft Word 2007 is no longer supported.`n`n`t`tScript will end.`n`n"
        AbortScript
    }
    Else {
        $ErrorActionPreference = $SaveEAPreference
        Write-Error "`n`n`t`tYou are running an untested or unsupported version of Microsoft Word.`n`n`t`tScript will end.`n`n`t`tPlease send info on your version of Word to webster@carlwebster.com`n`n"
        AbortScript
    }

    #only validate CompanyName if the field is blank
    If ([String]::IsNullOrEmpty($Script:CoName)) {
        Write-Verbose "$(Get-Date): Company name is blank.  Retrieve company name from registry."
        $TmpName = ValidateCompanyName

        If ([String]::IsNullOrEmpty($TmpName)) {
            Write-Warning "`n`n`t`tCompany Name is blank so Cover Page will not show a Company Name."
            Write-Warning "`n`t`tCheck HKCU:\Software\Microsoft\Office\Common\UserInfo for Company or CompanyName value."
            Write-Warning "`n`t`tYou may want to use the -CompanyName parameter if you need a Company Name on the cover page.`n`n"
        }
        Else {
            $Script:CoName = $TmpName
            Write-Verbose "$(Get-Date): Updated company name to $($Script:CoName)"
        }
    }

    If ($Script:WordCultureCode -ne "en-") {
        Write-Verbose "$(Get-Date): Check Default Cover Page for $($WordCultureCode)"
        [bool]$CPChanged = $False
        Switch ($Script:WordCultureCode) {
            'ca-'	{
                If ($CoverPage -eq "Sideline") {
                    $CoverPage = "Línia lateral"
                    $CPChanged = $True
                }
            }

            'da-'	{
                If ($CoverPage -eq "Sideline") {
                    $CoverPage = "Sidelinje"
                    $CPChanged = $True
                }
            }

            'de-'	{
                If ($CoverPage -eq "Sideline") {
                    $CoverPage = "Randlinie"
                    $CPChanged = $True
                }
            }

            'es-'	{
                If ($CoverPage -eq "Sideline") {
                    $CoverPage = "Línea lateral"
                    $CPChanged = $True
                }
            }

            'fi-'	{
                If ($CoverPage -eq "Sideline") {
                    $CoverPage = "Sivussa"
                    $CPChanged = $True
                }
            }

            'fr-'	{
                If ($CoverPage -eq "Sideline") {
                    If ($Script:WordVersion -eq $wdWord2013 -or $Script:WordVersion -eq $wdWord2016) {
                        $CoverPage = "Lignes latérales"
                        $CPChanged = $True
                    }
                    Else {
                        $CoverPage = "Ligne latérale"
                        $CPChanged = $True
                    }
                }
            }

            'nb-'	{
                If ($CoverPage -eq "Sideline") {
                    $CoverPage = "Sidelinje"
                    $CPChanged = $True
                }
            }

            'nl-'	{
                If ($CoverPage -eq "Sideline") {
                    $CoverPage = "Terzijde"
                    $CPChanged = $True
                }
            }

            'pt-'	{
                If ($CoverPage -eq "Sideline") {
                    $CoverPage = "Linha Lateral"
                    $CPChanged = $True
                }
            }

            'sv-'	{
                If ($CoverPage -eq "Sideline") {
                    $CoverPage = "Sidlinje"
                    $CPChanged = $True
                }
            }

            'zh-'	{
                If ($CoverPage -eq "Sideline") {
                    $CoverPage = "边线型"
                    $CPChanged = $True
                }
            }
        }

        If ($CPChanged) {
            Write-Verbose "$(Get-Date): Changed Default Cover Page from Sideline to $($CoverPage)"
        }
    }

    Write-Verbose "$(Get-Date): Validate cover page $($CoverPage) for culture code $($Script:WordCultureCode)"
    [bool]$ValidCP = $False

    $ValidCP = ValidateCoverPage $Script:WordVersion $CoverPage $Script:WordCultureCode

    If (!$ValidCP) {
        $ErrorActionPreference = $SaveEAPreference
        Write-Verbose "$(Get-Date): Word language value $($Script:WordLanguageValue)"
        Write-Verbose "$(Get-Date): Culture code $($Script:WordCultureCode)"
        Write-Error "`n`n`t`tFor $($Script:WordProduct), $($CoverPage) is not a valid Cover Page option.`n`n`t`tScript cannot continue.`n`n"
        AbortScript
    }

    ShowScriptOptions

    $Script:Word.Visible = $False

    #http://jdhitsolutions.com/blog/2012/05/san-diego-2012-powershell-deep-dive-slides-and-demos/
    #using Jeff's Demo-WordReport.ps1 file for examples
    Write-Verbose "$(Get-Date): Load Word Templates"

    [bool]$Script:CoverPagesExist = $False
    [bool]$BuildingBlocksExist = $False

    $Script:Word.Templates.LoadBuildingBlocks()
    #word 2010/2013/2016
    $BuildingBlocksCollection = $Script:Word.Templates | Where {$_.name -eq "Built-In Building Blocks.dotx"}

    Write-Verbose "$(Get-Date): Attempt to load cover page $($CoverPage)"
    $part = $Null

    $BuildingBlocksCollection |
        ForEach {
        If ($_.BuildingBlockEntries.Item($CoverPage).Name -eq $CoverPage) {
            $BuildingBlocks = $_
        }
    }

    If ($Null -ne $BuildingBlocks) {
        $BuildingBlocksExist = $True

        Try {
            $part = $BuildingBlocks.BuildingBlockEntries.Item($CoverPage)
        }

        Catch {
            $part = $Null
        }

        If ($Null -ne $part) {
            $Script:CoverPagesExist = $True
        }
    }

    If (!$Script:CoverPagesExist) {
        Write-Verbose "$(Get-Date): Cover Pages are not installed or the Cover Page $($CoverPage) does not exist."
        Write-Warning "Cover Pages are not installed or the Cover Page $($CoverPage) does not exist."
        Write-Warning "This report will not have a Cover Page."
    }

    Write-Verbose "$(Get-Date): Create empty word doc"
    $Script:Doc = $Script:Word.Documents.Add()
    If ($Null -eq $Script:Doc) {
        Write-Verbose "$(Get-Date): "
        $ErrorActionPreference = $SaveEAPreference
        Write-Error "`n`n`t`tAn empty Word document could not be created.`n`n`t`tScript cannot continue.`n`n"
        AbortScript
    }

    $Script:Selection = $Script:Word.Selection
    If ($Null -eq $Script:Selection) {
        Write-Verbose "$(Get-Date): "
        $ErrorActionPreference = $SaveEAPreference
        Write-Error "`n`n`t`tAn unknown error happened selecting the entire Word document for default formatting options.`n`n`t`tScript cannot continue.`n`n"
        AbortScript
    }

    #set Default tab stops to 1/2 inch (this line is not from Jeff Hicks)
    #36 = .50"
    $Script:Word.ActiveDocument.DefaultTabStop = 36

    #Disable Spell and Grammar Check to resolve issue and improve performance (from Pat Coughlin)
    Write-Verbose "$(Get-Date): Disable grammar and spell checking"
    #bug reported 1-Apr-2014 by Tim Mangan
    #save current options first before turning them off
    $Script:CurrentGrammarOption = $Script:Word.Options.CheckGrammarAsYouType
    $Script:CurrentSpellingOption = $Script:Word.Options.CheckSpellingAsYouType
    $Script:Word.Options.CheckGrammarAsYouType = $False
    $Script:Word.Options.CheckSpellingAsYouType = $False

    If ($BuildingBlocksExist) {
        #insert new page, getting ready for table of contents
        Write-Verbose "$(Get-Date): Insert new page, getting ready for table of contents"
        $part.Insert($Script:Selection.Range, $True) | Out-Null
        $Script:Selection.InsertNewPage()

        #table of contents
        Write-Verbose "$(Get-Date): Table of Contents - $($Script:MyHash.Word_TableOfContents)"
        $toc = $BuildingBlocks.BuildingBlockEntries.Item($Script:MyHash.Word_TableOfContents)
        If ($Null -eq $toc) {
            Write-Verbose "$(Get-Date): "
            Write-Verbose "$(Get-Date): Table of Content - $($Script:MyHash.Word_TableOfContents) could not be retrieved."
            Write-Warning "This report will not have a Table of Contents."
        }
        Else {
            $toc.insert($Script:Selection.Range, $True) | Out-Null
        }
    }
    Else {
        Write-Verbose "$(Get-Date): Table of Contents are not installed."
        Write-Warning "Table of Contents are not installed so this report will not have a Table of Contents."
    }

    #set the footer
    Write-Verbose "$(Get-Date): Set the footer"
    [string]$footertext = "Report created by $username"

    #get the footer
    Write-Verbose "$(Get-Date): Get the footer and format font"
    $Script:Doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekPrimaryFooter
    #get the footer and format font
    $footers = $Script:Doc.Sections.Last.Footers
    ForEach ($footer in $footers) {
        If ($footer.exists) {
            $footer.range.Font.name = "Calibri"
            $footer.range.Font.size = 8
            $footer.range.Font.Italic = $True
            $footer.range.Font.Bold = $True
        }
    } #end ForEach
    Write-Verbose "$(Get-Date): Footer text"
    $Script:Selection.HeaderFooter.Range.Text = $footerText

    #add page numbering
    Write-Verbose "$(Get-Date): Add page numbering"
    $Script:Selection.HeaderFooter.PageNumbers.Add($wdAlignPageNumberRight) | Out-Null

    FindWordDocumentEnd
    Write-Verbose "$(Get-Date):"
    #end of Jeff Hicks
}

Function UpdateDocumentProperties {
    Param([string]$AbstractTitle, [string]$SubjectTitle)
    #Update document properties
    If ($MSWORD -or $PDF) {
        If ($Script:CoverPagesExist) {
            Write-Verbose "$(Get-Date): Set Cover Page Properties"
            _SetDocumentProperty $Script:Doc.BuiltInDocumentProperties "Company" $Script:CoName
            _SetDocumentProperty $Script:Doc.BuiltInDocumentProperties "Title" $Script:title
            _SetDocumentProperty $Script:Doc.BuiltInDocumentProperties "Author" $username

            _SetDocumentProperty $Script:Doc.BuiltInDocumentProperties "Subject" $SubjectTitle

            #Get the Coverpage XML part
            $cp = $Script:Doc.CustomXMLParts | Where {$_.NamespaceURI -match "coverPageProps$"}

            #get the abstract XML part
            $ab = $cp.documentelement.ChildNodes | Where {$_.basename -eq "Abstract"}

            #set the text
            If ([String]::IsNullOrEmpty($Script:CoName)) {
                [string]$abstract = $AbstractTitle
            }
            Else {
                [string]$abstract = "$($AbstractTitle) for $($Script:CoName)"
            }

            $ab.Text = $abstract

            $ab = $cp.documentelement.ChildNodes | Where {$_.basename -eq "PublishDate"}
            #set the text
            [string]$abstract = (Get-Date -Format d).ToString()
            $ab.Text = $abstract

            Write-Verbose "$(Get-Date): Update the Table of Contents"
            #update the Table of Contents
            $Script:Doc.TablesOfContents.item(1).Update()
            $cp = $Null
            $ab = $Null
            $abstract = $Null
        }
    }
}

Function AddWordTable {
    #region Iain's Word table functions

    <#
        .SYNOPSIS
            Add a table to a Microsoft Word document
        .DESCRIPTION
            This function adds a table to a Microsoft Word document from either an array of
            Hashtables or an array of PSCustomObjects.

            Using this function is quicker than setting each table cell individually but can
            only utilise the built-in MS Word table autoformats. Individual tables cells can
            be altered after the table has been appended to the document (a table reference
            is returned).
        .EXAMPLE
            AddWordTable -Hashtable $HashtableArray

            This example adds table to the MS Word document, utilising all key/value pairs in
            the array of hashtables. Column headers will display the key names as defined.
            Note: the columns might not be displayed in the order that they were defined. To
            ensure columns are displayed in the required order utilise the -Columns parameter.
        .EXAMPLE
            AddWordTable -Hashtable $HashtableArray -List

            This example adds table to the MS Word document, utilising all key/value pairs in
            the array of hashtables. No column headers will be added, in a ListView format.
            Note: the columns might not be displayed in the order that they were defined. To
            ensure columns are displayed in the required order utilise the -Columns parameter.
        .EXAMPLE
            AddWordTable -CustomObject $PSCustomObjectArray

            This example adds table to the MS Word document, utilising all note property names
            the array of PSCustomObjects. Column headers will display the note property names.
            Note: the columns might not be displayed in the order that they were defined. To
            ensure columns are displayed in the required order utilise the -Columns parameter.
        .EXAMPLE
            AddWordTable -Hashtable $HashtableArray -Columns FirstName,LastName,EmailAddress

            This example adds a table to the MS Word document, but only using the specified
            key names: FirstName, LastName and EmailAddress. If other keys are present in the
            array of Hashtables they will be ignored.
        .EXAMPLE
            AddWordTable -CustomObject $PSCustomObjectArray -Columns FirstName,LastName,EmailAddress -Headers "First Name","Last Name","Email Address"

            This example adds a table to the MS Word document, but only using the specified
            PSCustomObject note properties: FirstName, LastName and EmailAddress. If other note
            properties are present in the array of PSCustomObjects they will be ignored. The
            display names for each specified column header has been overridden to display a
            custom header. Note: the order of the header names must match the specified columns.
    #>

        [CmdletBinding()]
        Param
        (
            # Array of Hashtable (including table headers)
            [Parameter(
                Mandatory = $True,
                ValueFromPipelineByPropertyName = $True,
                ParameterSetName = 'Hashtable',
                Position = 0
            )]
            [ValidateNotNullOrEmpty()]
            [System.Collections.Hashtable[]] $Hashtable,

            # Array of PSCustomObjects
            [Parameter(
                Mandatory = $True,
                ValueFromPipelineByPropertyName = $True,
                ParameterSetName = 'CustomObject',
                Position = 1
            )]
            [ValidateNotNullOrEmpty()]
            [PSCustomObject[]] $CustomObject,

            # Array of Hashtable key names or PSCustomObject property names to include, in display order.
            # If not supplied then all Hashtable keys or all PSCustomObject properties will be displayed.
            [Parameter(
                ValueFromPipelineByPropertyName = $True
            )]
            [AllowNull()]
            [string[]] $Columns = $Null,

            # Array of custom table header strings in display order.
            [Parameter(
                ValueFromPipelineByPropertyName = $True
            )]
            [AllowNull()]
            [string[]] $Headers = $Null,

            # AutoFit table behavior.
            [Parameter(
                ValueFromPipelineByPropertyName = $True
            )]
            [AllowNull()]
            [int] $AutoFit = -1,

            # List view (no headers)
            [Switch] $List,

            # Grid lines
            [Switch] $NoGridLines,

            [Switch] $NoInternalGridLines,

            # Built-in Word table formatting style constant
            # Would recommend only $wdTableFormatContempory for normal usage (possibly $wdTableFormatList5 for List view)
            [Parameter(
                ValueFromPipelineByPropertyName = $True
            )]
            [int] $Format = 0
        )

        Begin {
            Write-Debug ("Using parameter set '{0}'" -f $PSCmdlet.ParameterSetName);
            ## Check if -Columns wasn't specified but -Headers were (saves some additional parameter sets!)
            If (($Columns -eq $Null) -and ($Headers -ne $Null)) {
                Write-Warning "No columns specified and therefore, specified headers will be ignored.";
                $Columns = $Null;
            }
            ElseIf (($Columns -ne $Null) -and ($Headers -ne $Null)) {
                ## Check if number of specified -Columns matches number of specified -Headers
                If ($Columns.Length -ne $Headers.Length) {
                    Write-Error "The specified number of columns does not match the specified number of headers.";
                }
            } ## end elseif
        } ## end Begin

        Process {
            ## Build the Word table data string to be converted to a range and then a table later.
            [System.Text.StringBuilder] $WordRangeString = New-Object System.Text.StringBuilder;

            Switch ($PSCmdlet.ParameterSetName) {
                'CustomObject' {
                    If ($Columns -eq $Null) {
                        ## Build the available columns from all availble PSCustomObject note properties
                        [string[]] $Columns = @();
                        ## Add each NoteProperty name to the array
                        ForEach ($Property in ($CustomObject | Get-Member -MemberType NoteProperty)) {
                            $Columns += $Property.Name;
                        }
                    }

                    ## Add the table headers from -Headers or -Columns (except when in -List(view)
                    If (-not $List) {
                        Write-Debug ("$(Get-Date): `t`tBuilding table headers");
                        If ($Headers -ne $Null) {
                            [ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Headers));
                        }
                        Else {
                            [ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Columns));
                        }
                    }

                    ## Iterate through each PSCustomObject
                    Write-Debug ("$(Get-Date): `t`tBuilding table rows");
                    ForEach ($Object in $CustomObject) {
                        $OrderedValues = @();
                        ## Add each row item in the specified order
                        ForEach ($Column in $Columns) {
                            $OrderedValues += $Object.$Column;
                        }
                        ## Use the ordered list to add each column in specified order
                        [ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $OrderedValues));
                    } ## end foreach
                    Write-Debug ("$(Get-Date): `t`t`tAdded '{0}' table rows" -f ($CustomObject.Count));
                } ## end CustomObject

                Default {
                    ## Hashtable
                    If ($Columns -eq $Null) {
                        ## Build the available columns from all available hashtable keys. Hopefully
                        ## all Hashtables have the same keys (they should for a table).
                        $Columns = $Hashtable[0].Keys;
                    }

                    ## Add the table headers from -Headers or -Columns (except when in -List(view)
                    If (-not $List) {
                        Write-Debug ("$(Get-Date): `t`tBuilding table headers");
                        If ($Headers -ne $Null) {
                            [ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Headers));
                        }
                        Else {
                            [ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Columns));
                        }
                    }

                    ## Iterate through each Hashtable
                    Write-Debug ("$(Get-Date): `t`tBuilding table rows");
                    ForEach ($Hash in $Hashtable) {
                        $OrderedValues = @();
                        ## Add each row item in the specified order
                        ForEach ($Column in $Columns) {
                            $OrderedValues += $Hash.$Column;
                        }
                        ## Use the ordered list to add each column in specified order
                        [ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $OrderedValues));
                    } ## end foreach

                    Write-Debug ("$(Get-Date): `t`t`tAdded '{0}' table rows" -f $Hashtable.Count);
                } ## end default
            } ## end switch

            ## Create a MS Word range and set its text to our tab-delimited, concatenated string
            Write-Debug ("$(Get-Date): `t`tBuilding table range");
            $WordRange = $Script:Doc.Application.Selection.Range;
            $WordRange.Text = $WordRangeString.ToString();

            ## Create hash table of named arguments to pass to the ConvertToTable method
            $ConvertToTableArguments = @{ Separator = [Microsoft.Office.Interop.Word.WdTableFieldSeparator]::wdSeparateByTabs; }

            ## Negative built-in styles are not supported by the ConvertToTable method
            If ($Format -ge 0) {
                $ConvertToTableArguments.Add("Format", $Format);
                $ConvertToTableArguments.Add("ApplyBorders", $True);
                $ConvertToTableArguments.Add("ApplyShading", $True);
                $ConvertToTableArguments.Add("ApplyFont", $True);
                $ConvertToTableArguments.Add("ApplyColor", $True);
                If (!$List) {
                    $ConvertToTableArguments.Add("ApplyHeadingRows", $True);
                }
                $ConvertToTableArguments.Add("ApplyLastRow", $True);
                $ConvertToTableArguments.Add("ApplyFirstColumn", $True);
                $ConvertToTableArguments.Add("ApplyLastColumn", $True);
            }

            ## Invoke ConvertToTable method - with named arguments - to convert Word range to a table
            ## See http://msdn.microsoft.com/en-us/library/office/aa171893(v=office.11).aspx
            Write-Debug ("$(Get-Date): `t`tConverting range to table");
            ## Store the table reference just in case we need to set alternate row coloring
            $WordTable = $WordRange.GetType().InvokeMember(
                "ConvertToTable", # Method name
                [System.Reflection.BindingFlags]::InvokeMethod, # Flags
                $Null, # Binder
                $WordRange, # Target (self!)
                ([Object[]]($ConvertToTableArguments.Values)), ## Named argument values
                $Null, # Modifiers
                $Null, # Culture
                ([String[]]($ConvertToTableArguments.Keys))     ## Named argument names
            );

            ## Implement grid lines (will wipe out any existing formatting
            If ($Format -lt 0) {
            Write-Debug ("$(Get-Date): `t`tSetting table format");
                $WordTable.Style = $Format;
            }

            ## Set the table autofit behavior
            If ($AutoFit -ne -1) {
                $WordTable.AutoFitBehavior($AutoFit);
            }

            If (!$List) {
                #the next line causes the heading row to flow across page breaks
                $WordTable.Rows.First.Headingformat = $wdHeadingFormatTrue;
            }

            If (!$NoGridLines) {
                $WordTable.Borders.InsideLineStyle = $wdLineStyleSingle;
                $WordTable.Borders.OutsideLineStyle = $wdLineStyleSingle;
            }
            If ($NoGridLines) {
                $WordTable.Borders.InsideLineStyle = $wdLineStyleNone;
                $WordTable.Borders.OutsideLineStyle = $wdLineStyleNone;
            }
            If ($NoInternalGridLines) {
                $WordTable.Borders.InsideLineStyle = $wdLineStyleNone;
                $WordTable.Borders.OutsideLineStyle = $wdLineStyleSingle;
            }

            Return $WordTable;

        } ## end Process
}

Function SetWordCellFormat {
    <#
    .Synopsis
        Sets the format of one or more Word table cells
    .DESCRIPTION
        This function sets the format of one or more table cells, either from a collection
        of Word COM object cell references, an individual Word COM object cell reference or
        a hashtable containing Row and Column information.

        The font name, font size, bold, italic , underline and shading values can be used.
    .EXAMPLE
        SetWordCellFormat -Hashtable $Coordinates -Table $TableReference -Bold

        This example sets all text to bold that is contained within the $TableReference
        Word table, using an array of hashtables. Each hashtable contain a pair of co-
        ordinates that is used to select the required cells. Note: the hashtable must
        contain the .Row and .Column key names. For example:
        @ { Row = 7; Column = 3 } to set the cell at row 7 and column 3 to bold.
    .EXAMPLE
        $RowCollection = $Table.Rows.First.Cells
        SetWordCellFormat -Collection $RowCollection -Bold -Size 10

        This example sets all text to size 8 and bold for all cells that are contained
        within the first row of the table.
        Note: the $Table.Rows.First.Cells returns a collection of Word COM cells objects
        that are in the first table row.
    .EXAMPLE
        $ColumnCollection = $Table.Columns.Item(2).Cells
        SetWordCellFormat -Collection $ColumnCollection -BackgroundColor 255

        This example sets the background (shading) of all cells in the table's second
        column to red.
        Note: the $Table.Columns.Item(2).Cells returns a collection of Word COM cells objects
        that are in the table's second column.
    .EXAMPLE
        SetWordCellFormat -Cell $Table.Cell(17,3) -Font "Tahoma" -Color 16711680

        This example sets the font to Tahoma and the text color to blue for the cell located
        in the table's 17th row and 3rd column.
        Note: the $Table.Cell(17,3) returns a single Word COM cells object.
    #>
    [CmdletBinding(DefaultParameterSetName = 'Collection')]
    Param (
        # Word COM object cell collection reference
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ParameterSetName = 'Collection', Position = 0)] [ValidateNotNullOrEmpty()] $Collection,
        # Word COM object individual cell reference
        [Parameter(Mandatory = $true, ParameterSetName = 'Cell', Position = 0)] [ValidateNotNullOrEmpty()] $Cell,
        # Hashtable of cell co-ordinates
        [Parameter(Mandatory = $true, ParameterSetName = 'Hashtable', Position = 0)] [ValidateNotNullOrEmpty()] [System.Collections.Hashtable[]] $Coordinates,
        # Word COM object table reference
        [Parameter(Mandatory = $true, ParameterSetName = 'Hashtable', Position = 1)] [ValidateNotNullOrEmpty()] $Table,
        # Font name
        [Parameter()] [AllowNull()] [string] $Font = $null,
        # Font color
        [Parameter()] [AllowNull()] $Color = $null,
        # Font size
        [Parameter()] [ValidateNotNullOrEmpty()] [int] $Size = 0,
        # Cell background color
        [Parameter()] [AllowNull()] $BackgroundColor = $null,
        # Force solid background color
        [Switch] $Solid,
        [Switch] $Bold,
        [Switch] $Italic,
        [Switch] $Underline
    )

    Begin {
        Write-Debug ("Using parameter set '{0}'." -f $PSCmdlet.ParameterSetName);
    }

    Process {
        Switch ($PSCmdlet.ParameterSetName) {
            'Collection' {
                ForEach ($Cell in $Collection) {
                    If ($BackgroundColor -ne $null) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
                    If ($Bold) { $Cell.Range.Font.Bold = $true; }
                    If ($Italic) { $Cell.Range.Font.Italic = $true; }
                    If ($Underline) { $Cell.Range.Font.Underline = 1; }
                    If ($Font -ne $null) { $Cell.Range.Font.Name = $Font; }
                    If ($Color -ne $null) { $Cell.Range.Font.Color = $Color; }
                    If ($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
                    If ($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
                } # end foreach
            } # end Collection
            'Cell' {
                If ($Bold) { $Cell.Range.Font.Bold = $true; }
                If ($Italic) { $Cell.Range.Font.Italic = $true; }
                If ($Underline) { $Cell.Range.Font.Underline = 1; }
                If ($Font -ne $null) { $Cell.Range.Font.Name = $Font; }
                If ($Color -ne $null) { $Cell.Range.Font.Color = $Color; }
                If ($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
                If ($BackgroundColor -ne $null) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
                If ($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
            } # end Cell
            'Hashtable' {
                ForEach ($Coordinate in $Coordinates) {
                    $Cell = $Table.Cell($Coordinate.Row, $Coordinate.Column);
                    If ($Bold) { $Cell.Range.Font.Bold = $true; }
                    If ($Italic) { $Cell.Range.Font.Italic = $true; }
                    If ($Underline) { $Cell.Range.Font.Underline = 1; }
                    If ($Font -ne $null) { $Cell.Range.Font.Name = $Font; }
                    If ($Color -ne $null) { $Cell.Range.Font.Color = $Color; }
                    If ($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
                    If ($BackgroundColor -ne $null) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
                    If ($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
                }
            } # end Hashtable
        } # end switch
    } # end process
}

Function SetWordTableAlternateRowColor {
    <#
    .Synopsis
        Sets alternate row colors in a Word table
    .DESCRIPTION
        This function sets the format of alternate rows within a Word table using the
        specified $BackgroundColor. This function is expensive (in performance terms) as
        it recursively sets the format on alternate rows. It would be better to pick one
        of the predefined table formats (if one exists)? Obviously the more rows, the
        longer it takes :'(

        Note: this function is called by the AddWordTable function if an alternate row
        format is specified.
    .EXAMPLE
        SetWordTableAlternateRowColor -Table $TableReference -BackgroundColor 255

        This example sets every-other table (starting with the first) row and sets the
        background color to red (wdColorRed).
    .EXAMPLE
        SetWordTableAlternateRowColor -Table $TableReference -BackgroundColor 39423 -Seed Second

        This example sets every other table (starting with the second) row and sets the
        background color to light orange (weColorLightOrange).
    #>
    [CmdletBinding()]
    Param (
        # Word COM object table reference
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, Position = 0)] [ValidateNotNullOrEmpty()] $Table,
        # Alternate row background color
        [Parameter(Mandatory = $true, Position = 1)] [ValidateNotNull()] [int] $BackgroundColor,
        # Alternate row starting seed
        [Parameter(ValueFromPipelineByPropertyName = $true, Position = 2)] [ValidateSet('First', 'Second')] [string] $Seed = 'First'
    )

    Process {
        $StartDateTime = Get-Date;
        Write-Debug ("{0}: `t`tSetting alternate table row colors.." -f $StartDateTime);

        ## Determine the row seed (only really need to check for 'Second' and default to 'First' otherwise
        If ($Seed.ToLower() -eq 'second') {
            $StartRowIndex = 2;
        }
        Else {
            $StartRowIndex = 1;
        }

        For ($AlternateRowIndex = $StartRowIndex; $AlternateRowIndex -lt $Table.Rows.Count; $AlternateRowIndex += 2) {
            $Table.Rows.Item($AlternateRowIndex).Shading.BackgroundPatternColor = $BackgroundColor;
        }

        ## I've put verbose calls in here we can see how expensive this functionality actually is.
        $EndDateTime = Get-Date;
        $ExecutionTime = New-TimeSpan -Start $StartDateTime -End $EndDateTime;
        Write-Debug ("{0}: `t`tDone setting alternate row style color in '{1}' seconds" -f $EndDateTime, $ExecutionTime.TotalSeconds);
    }
}
#endregion

#region registry functions
#http://stackoverflow.com/questions/5648931/test-if-registry-value-exists
# This Function just gets $True or $False
Function Test-RegistryValue($path, $name) {
    $key = Get-Item -LiteralPath $path -EA 0
    $key -and $Null -ne $key.GetValue($name, $Null)
}

# Gets the specified registry value or $Null if it is missing
Function Get-RegistryValue($path, $name) {
    $key = Get-Item -LiteralPath $path -EA 0
    If ($key) {
        $key.GetValue($name, $Null)
    }
    Else {
        $Null
    }
}
#endregion



Function validStateProp( [object] $object, [string] $topLevel, [string] $secondLevel )
{
	#function created 8-jan-2015 by Michael B. Smith
	if( $object )
	{
		If( ( gm -Name $topLevel -InputObject $object ) )
		{
			If( ( gm -Name $secondLevel -InputObject $object.$topLevel ) )
			{
				Return $True
			}
		}
	}
	Return $False
}

#region general script functions
Function ShowScriptOptions {
    Write-Verbose "$(Get-Date): "
    Write-Verbose "$(Get-Date): "
    Write-Verbose "$(Get-Date): AddDateTime   : $($AddDateTime)"
    Write-Verbose "$(Get-Date): AdminAddress  : $($AdminAddress)"
    If ($MSWORD -or $PDF) {
        Write-Verbose "$(Get-Date): Company Name  : $($Script:CoName)"
        Write-Verbose "$(Get-Date): Cover Page    : $($CoverPage)"
    }
    Write-Verbose "$(Get-Date): Dev           : $($Dev)"
    If ($Dev) {
        Write-Verbose "$(Get-Date): DevErrorFile  : $($Script:DevErrorFile)"
    }
    Write-Verbose "$(Get-Date): Domain        : $($Domain)"
    Write-Verbose "$(Get-Date): End Date      : $($EndDate)"
    Write-Verbose "$(Get-Date): Filename1     : $($Script:filename1)"
    If ($PDF) {
        Write-Verbose "$(Get-Date): Filename2     : $($Script:filename2)"
    }
    Write-Verbose "$(Get-Date): Folder        : $($Folder)"
    Write-Verbose "$(Get-Date): From          : $($From)"
    Write-Verbose "$(Get-Date): HW Inventory  : $($Hardware)"
    Write-Verbose "$(Get-Date): Save As HTML  : $($HTML)"
    Write-Verbose "$(Get-Date): Save As PDF   : $($PDF)"
    Write-Verbose "$(Get-Date): Save As TEXT  : $($TEXT)"
    Write-Verbose "$(Get-Date): Save As WORD  : $($MSWORD)"
    Write-Verbose "$(Get-Date): ScriptInfo    : $($ScriptInfo)"
    Write-Verbose "$(Get-Date): Smtp Port     : $($SmtpPort)"
    Write-Verbose "$(Get-Date): Smtp Server   : $($SmtpServer)"
    Write-Verbose "$(Get-Date): Start Date    : $($StartDate)"
    Write-Verbose "$(Get-Date): Title         : $($Script:Title)"
    Write-Verbose "$(Get-Date): To            : $($To)"
    Write-Verbose "$(Get-Date): Use SSL       : $($UseSSL)"
    Write-Verbose "$(Get-Date): User          : $($User)"
    If ($MSWORD -or $PDF) {
        Write-Verbose "$(Get-Date): User Name     : $($UserName)"
    }
    Write-Verbose "$(Get-Date): "
    Write-Verbose "$(Get-Date): OS Detected   : $($Script:RunningOS)"
    Write-Verbose "$(Get-Date): PoSH version  : $($Host.Version)"
    Write-Verbose "$(Get-Date): PSCulture     : $($PSCulture)"
    Write-Verbose "$(Get-Date): PSUICulture   : $($PSUICulture)"
    If ($MSWORD -or $PDF) {
        Write-Verbose "$(Get-Date): Word language : $($Script:WordLanguageValue)"
        Write-Verbose "$(Get-Date): Word version  : $($Script:WordProduct)"
    }
    Write-Verbose "$(Get-Date): "
    Write-Verbose "$(Get-Date): Script start  : $($Script:StartTime)"
    Write-Verbose "$(Get-Date): "
    Write-Verbose "$(Get-Date): "
}

Function SaveandCloseDocumentandShutdownWord {
    #bug fix 1-Apr-2014
    #reset Grammar and Spelling options back to their original settings
    $Script:Word.Options.CheckGrammarAsYouType = $Script:CurrentGrammarOption
    $Script:Word.Options.CheckSpellingAsYouType = $Script:CurrentSpellingOption

    Write-Verbose "$(Get-Date): Save and Close document and Shutdown Word"
    If ($Script:WordVersion -eq $wdWord2010) {
        #the $saveFormat below passes StrictMode 2
        #I found this at the following two links
        #http://blogs.technet.com/b/bshukla/archive/2011/09/27/3347395.aspx
        #http://msdn.microsoft.com/en-us/library/microsoft.office.interop.word.wdsaveformat(v=office.14).aspx
        If ($PDF) {
            Write-Verbose "$(Get-Date): Saving as DOCX file first before saving to PDF"
        }
        Else {
            Write-Verbose "$(Get-Date): Saving DOCX file"
        }
        If ($AddDateTime) {
            $Script:FileName1 += "_$(Get-Date -f yyyy-MM-dd_HHmm).docx"
            If ($PDF) {
                $Script:FileName2 += "_$(Get-Date -f yyyy-MM-dd_HHmm).pdf"
            }
        }
        Write-Verbose "$(Get-Date): Running $($Script:WordProduct) and detected operating system $($Script:RunningOS)"
        $saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatDocumentDefault")
        $Script:Doc.SaveAs([REF]$Script:FileName1, [ref]$SaveFormat)
        If ($PDF) {
            Write-Verbose "$(Get-Date): Now saving as PDF"
            $saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatPDF")
            $Script:Doc.SaveAs([REF]$Script:FileName2, [ref]$saveFormat)
        }
    }
    ElseIf ($Script:WordVersion -eq $wdWord2013 -or $Script:WordVersion -eq $wdWord2016) {
        If ($PDF) {
            Write-Verbose "$(Get-Date): Saving as DOCX file first before saving to PDF"
        }
        Else {
            Write-Verbose "$(Get-Date): Saving DOCX file"
        }
        If ($AddDateTime) {
            $Script:FileName1 += "_$(Get-Date -f yyyy-MM-dd_HHmm).docx"
            If ($PDF) {
                $Script:FileName2 += "_$(Get-Date -f yyyy-MM-dd_HHmm).pdf"
            }
        }
        Write-Verbose "$(Get-Date): Running $($Script:WordProduct) and detected operating system $($Script:RunningOS)"
        $Script:Doc.SaveAs2([REF]$Script:FileName1, [ref]$wdFormatDocumentDefault)
        If ($PDF) {
            Write-Verbose "$(Get-Date): Now saving as PDF"
            $Script:Doc.SaveAs([REF]$Script:FileName2, [ref]$wdFormatPDF)
        }
    }

    Write-Verbose "$(Get-Date): Closing Word"
    $Script:Doc.Close()
    $Script:Word.Quit()
    If ($PDF) {
        [int]$cnt = 0
        While (Test-Path $Script:FileName1) {
            $cnt++
            If ($cnt -gt 1) {
                Write-Verbose "$(Get-Date): Waiting another 10 seconds to allow Word to fully close (try # $($cnt))"
                Start-Sleep -Seconds 10
                $Script:Word.Quit()
                If ($cnt -gt 2) {
                    #kill the winword process

                    #find out our session (usually "1" except on TS/RDC or Citrix)
                    $SessionID = (Get-Process -PID $PID).SessionId

                    #Find out if winword is running in our session
                    $wordprocess = ((Get-Process 'WinWord' -ea 0)|? {$_.SessionId -eq $SessionID}).Id
                    If ($wordprocess -gt 0) {
                        Write-Verbose "$(Get-Date): Attempting to stop WinWord process # $($wordprocess)"
                        Stop-Process $wordprocess -EA 0
                    }
                }
            }
            Write-Verbose "$(Get-Date): Attempting to delete $($Script:FileName1) since only $($Script:FileName2) is needed (try # $($cnt))"
            Remove-Item $Script:FileName1 -EA 0 4>$Null
        }
    }
    Write-Verbose "$(Get-Date): System Cleanup"
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Script:Word) | Out-Null
    If (Test-Path variable:global:word) {
        Remove-Variable -Name word -Scope Global 4>$Null
    }
    $SaveFormat = $Null
    [gc]::collect()
    [gc]::WaitForPendingFinalizers()

    #is the winword process still running? kill it

    #find out our session (usually "1" except on TS/RDC or Citrix)
    $SessionID = (Get-Process -PID $PID).SessionId

    #Find out if winword is running in our session
    $wordprocess = $Null
    $wordprocess = ((Get-Process 'WinWord' -ea 0)|? {$_.SessionId -eq $SessionID}).Id
    If ($null -ne $wordprocess -and $wordprocess -gt 0) {
        Write-Verbose "$(Get-Date): WinWord process is still running. Attempting to stop WinWord process # $($wordprocess)"
        Stop-Process $wordprocess -EA 0
    }
}

Function UpdateDocumentProperties
{
	Param([string]$AbstractTitle, [string]$SubjectTitle)
	#Update document properties
	If($MSWORD -or $PDF)
	{
        If ($Script:CoverPagesExist) {
            Write-Verbose "$(Get-Date): Set Cover Page Properties"
            #_SetDocumentProperty $Script:Doc.BuiltInDocumentProperties "Company" $Script:CoName
            #_SetDocumentProperty $Script:Doc.BuiltInDocumentProperties "Title" $Script:title
			#_SetDocumentProperty $Script:Doc.BuiltInDocumentProperties "Author" $username

            #_SetDocumentProperty $Script:Doc.BuiltInDocumentProperties "Subject" $SubjectTitle

            Set-DocumentProperty -Document $Script:Doc -DocProperty Title -Value $Script:title
            Set-DocumentProperty -Document $Script:Doc -DocProperty Company -Value $Script:CoName
            Set-DocumentProperty -Document $Script:Doc -DocProperty Author -Value $UserName
            Set-DocumentProperty -Document $Script:Doc -DocProperty Subject -Value $SubjectTitle

            #Get the Coverpage XML part
            $cp = $Script:Doc.CustomXMLParts | Where {$_.NamespaceURI -match "coverPageProps$"}

            #get the abstract XML part
            $ab = $cp.documentelement.ChildNodes | Where {$_.basename -eq "Abstract"}

            #set the text
            If ([String]::IsNullOrEmpty($Script:CoName)) {
                [string]$abstract = $AbstractTitle
            }
            Else {
                [string]$abstract = "$($AbstractTitle) for $($Script:CoName)"
            }

            $ab.Text = $abstract

            $ab = $cp.documentelement.ChildNodes | Where {$_.basename -eq "PublishDate"}
            #set the text
            [string]$abstract = (Get-Date -Format d).ToString()
            $ab.Text = $abstract

            Write-Verbose "$(Get-Date): Update the Table of Contents"
            #update the Table of Contents
            $Script:Doc.TablesOfContents.item(1).Update()
            $cp = $Null
            $ab = $Null
            $abstract = $Null
        }
	}
}

. .\Set-WordLine.ps1

. .\Recursive.ps1

. .\Set-DocumentProperty.ps1

Get-Process -Name *Word* | Stop-Process

Start-Sleep 2

CheckWordPreReq

SetupWord

$PVSdata = Get-Content (Join-Path $PSScriptRoot pvs.json) | ConvertFrom-Json

try {
    $PVSData | Convert-ObjToDoc | Where-Object {$_.LineType -eq 'Heading'} | Set-WordHeadingLine
}
catch {
    Write-host 'bug'
}

UpdateDocumentProperties

SaveandCloseDocumentandShutdownWord