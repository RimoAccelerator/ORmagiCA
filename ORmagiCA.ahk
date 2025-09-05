#Requires AutoHotkey v2.0

; Global configuration
global OFAKE_G_PATH := "D:\CompChem\OfakeG.exe"
global GAUSS_VIEW_PATH := "D:\GaussView\GV6.0.16WIN\g16w\gview.exe"
global SETTINGS_GUI := ""
global CURRENT_KEYWORDS := "No change"

; Initialize on startup
InitializeSettings()

; Initialize settings from INI file
InitializeSettings() {
    global OFAKE_G_PATH, GAUSS_VIEW_PATH, CURRENT_KEYWORDS, SETTINGS_GUI
    
    iniPath := A_ScriptDir "\ORmagiCA_settings.ini"
    
    ; Read paths from ini if exists, otherwise keep defaults
    if FileExist(iniPath) {
        savedGV := IniRead(iniPath, "Paths", "GaussView", "")
        savedOF := IniRead(iniPath, "Paths", "OfakeG", "")
        
        if (savedGV != "")
            GAUSS_VIEW_PATH := savedGV
        if (savedOF != "")
            OFAKE_G_PATH := savedOF
    }
    
    ; Create settings GUI
    if !IsObject(SETTINGS_GUI)
        SETTINGS_GUI := SettingsGui()
}

; Settings GUI Class
class SettingsGui {
    __New() {
        ; Create window with shadow style (CS_DROPSHADOW = 0x20000)
        this.gui := Gui("+AlwaysOnTop -Caption +ToolWindow +LastFound")
        DllCall("SetClassLong", "Ptr", this.gui.Hwnd, "Int", -26, "Int", DllCall("GetClassLong", "Ptr", this.gui.Hwnd, "Int", -26) | 0x20000)
        
        this.gui.SetFont("s10", "Segoe UI")
        this.gui.BackColor := "FFFFFF"
        
        ; Create controls
        this.gui.Add("Text", "x10 y10", "GaussView Path:")
        this.gvPath := this.gui.Add("Edit", "x10 y30 w400", GAUSS_VIEW_PATH)
        
        this.gui.Add("Text", "x10 y60", "OfakeG Path:")
        this.ofakePath := this.gui.Add("Edit", "x10 y80 w400", OFAKE_G_PATH)
        
        ; ORCA Keywords section with "No change" as protected item
        this.gui.Add("Text", "x10 y110", "ORCA Keywords:")
        this.keywordsList := this.gui.Add("ListBox", "x10 y130 w400 h150 vKeywordsList")
        
        ; Add keyword input and buttons
        this.newKeyword := this.gui.Add("Edit", "x10 y290 w300")
        addBtn := this.gui.Add("Button", "x320 y290 w90 h25", "Add")
        deleteBtn := this.gui.Add("Button", "x320 y320 w90 h25", "Delete")
        
        ; Events
        this.gvPath.OnEvent("Change", this.SavePaths.Bind(this))
        this.ofakePath.OnEvent("Change", this.SavePaths.Bind(this))
        this.keywordsList.OnEvent("Change", this.UpdateKeywords.Bind(this))
        this.keywordsList.OnEvent("Change", this.OnKeywordSelected.Bind(this))  ; Add this line
        addBtn.OnEvent("Click", this.AddKeyword.Bind(this))
        deleteBtn.OnEvent("Click", this.DeleteKeyword.Bind(this))
        
        ; Handle Enter key in new keyword edit box
        this.newKeyword.OnEvent("Change", this.OnNewKeywordChange.Bind(this))
        
        ; Handle Escape key and focus loss
        this.gui.OnEvent("Escape", this.Hide.Bind(this))
        
        ; Setup timer to check focus
        SetTimer(this.CheckFocus.Bind(this), 100)
        
        ; Add default keywords
        this.LoadKeywords()
        
        return this
    }

    Show() {
        ; Get primary monitor's work area (excludes taskbar)
        MonitorGetWorkArea(MonitorGetPrimary(), &monLeft, &monTop, &monRight, &monBottom)
        screenWidth := monRight - monLeft
        screenHeight := monBottom - monTop
        
        ; Fixed window dimensions
        winWidth := 420
        winHeight := 360
        
        ; Calculate center position (accounting for monitor position)
        x := monLeft + (screenWidth - winWidth) / 2
        y := monTop + (screenHeight - winHeight) / 2
        
        ; Show window at center with shadow
        this.gui.Show(Format("x{} y{} w{} h{}", Round(x), Round(y), winWidth, winHeight))
    }
    
    OnNewKeywordChange(*) {
        if (GetKeyState("Enter")) {
            this.AddKeyword()
        }
    }
    
    AddKeyword(*) {
        newKeyword := this.newKeyword.Text
        if (newKeyword != "") {
            ; Check if keyword already exists to avoid duplicates
            for keyword in this.keywords {
                if (keyword = newKeyword) {
                    this.newKeyword.Value := ""  ; Clear input
                    return  ; Don't add duplicate
                }
            }
            
            this.keywords.Push(newKeyword)
            this.keywordsList.Add([newKeyword])
            this.newKeyword.Value := ""  ; Clear input
            this.SaveKeywords()
        }
    }
    
    DeleteKeyword(*) {
        selectedIndex := this.keywordsList.Value
        if (selectedIndex > 1) {  ; Don't delete "No change"
            selectedText := this.keywordsList.Text
            
            ; Find the actual index in the keywords array that matches the selected text
            for i, keyword in this.keywords {
                if (keyword = selectedText) {
                    this.keywords.RemoveAt(i)
                    break
                }
            }
            
            ; Delete from the ListBox
            this.keywordsList.Delete(selectedIndex)
            this.SaveKeywords()
        }
    }
    
    SaveKeywords() {
        iniPath := A_ScriptDir "\ORmagiCA_settings.ini"
        
        ; Clear existing keywords
        IniDelete(iniPath, "Keywords")
        
        ; Save new keywords
        for i, keyword in this.keywords {
            if (i > 1)  ; Skip "No change"
                IniWrite(keyword, iniPath, "Keywords", "Keyword" (i-1))
        }
    }
    
    Hide(*) {
        this.gui.Hide()
    }
    
    CheckFocus() {
        ; Check if GUI handle exists and is visible
        if (!WinExist("ahk_id " this.gui.Hwnd))
            return
            
        activeHwnd := WinExist("A")
        if (activeHwnd != this.gui.Hwnd)
            this.Hide()
    }
    
    LoadKeywords() {
        this.keywordsList.Delete()
        this.keywords := ["No change"]  ; Initialize with "No change"
        
        iniPath := A_ScriptDir "\ORmagiCA_settings.ini"
        loop {
            keyword := IniRead(iniPath, "Keywords", "Keyword" A_Index, "")
            if (keyword = "")
                break
            this.keywords.Push(keyword)
        }
        
        ; Update ListBox
        for keyword in this.keywords
            this.keywordsList.Add([keyword])
    }
    
    SavePaths(*) {
        global OFAKE_G_PATH, GAUSS_VIEW_PATH
        
        iniPath := A_ScriptDir "\ORmagiCA_settings.ini"
        
        GAUSS_VIEW_PATH := this.gvPath.Text
        OFAKE_G_PATH := this.ofakePath.Text
        
        IniWrite(GAUSS_VIEW_PATH, iniPath, "Paths", "GaussView")
        IniWrite(OFAKE_G_PATH, iniPath, "Paths", "OfakeG")
    }
    
    UpdateKeywords(*) {
        global CURRENT_KEYWORDS
        
        if (this.keywordsList.Value)
            CURRENT_KEYWORDS := this.keywordsList.Text
    }
    
    OnKeywordSelected(*) {
        ; Update the text box with the selected keyword
        if (this.keywordsList.Value) {
            selectedText := this.keywordsList.Text
            this.newKeyword.Value := selectedText
        }
    }
}

; Hotkey definitions
; Global hotkey for file processing
^+g::
{
    ; Get currently selected files (now returns an array)
    selectedFiles := GetSelectedFiles()
    if (selectedFiles.Length = 0)
        return

    ; Process each selected file with a 100ms delay between them
    for i, selectedFile in selectedFiles
    {
        ; Get file extension
        SplitPath(selectedFile, &fileName, &filePath, &fileExt)
        SplitPath(selectedFile, &fileNameWithExt, &fileDir, &fileExt, &fileName, &drive)

        if (fileExt = "out")
            ProcessOrcaOutputFile(selectedFile, filePath, fileName)
        else if (fileExt = "inp")
        {
            fileContent := FileRead(selectedFile)
            keywords := ExtractKeywords(fileContent)
            nprocs := ExtractNprocs(fileContent)
            maxcore := ExtractMaxcore(fileContent)
            settings := ExtractSettings(fileContent, "inp")
            charge := ExtractCharge(fileContent)
            multiplicity := ExtractMultiplicity(fileContent)
            coordinates := ExtractCoordinates(fileContent)
            
            fakeGjfFile := filePath . "\" . fileName . "_fake.gjf"
            if CreateGaussianInput(fakeGjfFile, nprocs, maxcore, keywords, charge, multiplicity, coordinates, settings)
            {
                if FileExist(fakeGjfFile)
                {
                    OpenWithGaussView(fakeGjfFile)
                    SetTimer () => FileExist(fakeGjfFile) ? FileDelete(fakeGjfFile) : "", -2000
                }
            }
        }
        
        ; Add delay between processing files (unless this is the last file)
        if (i < selectedFiles.Length)
            Sleep(100)
    }
}

; GaussView specific hotkeys
#HotIf WinActive("ahk_exe gview.exe")
^+d::
{
    global SETTINGS_GUI
    if !IsObject(SETTINGS_GUI)
        SETTINGS_GUI := SettingsGui()
    SETTINGS_GUI.Show()
}
^+s::
{
    ; Store original window handle
    originalHwnd := WinExist("A")
    
    ; Send original Ctrl+S to trigger save dialog
    Send "^s"
    Sleep(200)
    
    ; Try to detect save dialog
    dialogHwnd := ""
    startTime := A_TickCount
    while (A_TickCount - startTime < 1000)
    {
        if (hwnd := WinExist("Save Structure Files"))
        {
            dialogHwnd := hwnd
            break
        }
        Sleep(100)
    }
    
    if (!dialogHwnd)
    {
        MsgBox("Save dialog not detected.", "Error", "Icon!")
        return
    }

    ; Wait for dialog to close
    WinWaitClose("ahk_id " dialogHwnd)
    Sleep(500)  ; Increased delay to ensure window title is updated
    
    ; Activate the main window and get the new title
    WinActivate("ahk_id " originalHwnd)
    Sleep(500)  ; Give time for window to activate
    
    currentTitle := WinGetTitle("ahk_id " originalHwnd)
    
    ; Extract path from current title
    currentPath := ""
    openParenPos := InStr(currentTitle, "(")
    closeParenPos := InStr(currentTitle, ")")
    
    if (openParenPos && closeParenPos && openParenPos < closeParenPos)
    {
        fullPath := SubStr(currentTitle, openParenPos + 1, closeParenPos - openParenPos - 1)
        fullPath := Trim(fullPath)
        currentPath := StrReplace(fullPath, "/", "\")
    }
    
    ; Process the path and create saved path
    savedPath := ""
    if (currentPath)
    {
        ; Convert gif to gjf
        savedPath := RegExReplace(currentPath, "\.gif$", ".gjf")
        savedPath := RegExReplace(savedPath, "\.gi[f]$", ".gjf")
    }
    
    if (!savedPath)
    {
        MsgBox("Unable to get saved file path.`nFull title: " currentTitle, "Error", "Icon!")
        return
    }
    
    ; Wait for file to exist
    startTime := A_TickCount
    while (!FileExist(savedPath) && A_TickCount - startTime < 1000)
        Sleep(100)
        
    if (!FileExist(savedPath))
    {
        MsgBox("File not successfully saved: " savedPath, "Error", "Icon!")
        return
    }

    ; Process the saved file
    try
    {
        fileContent := FileRead(savedPath)
        if (fileContent = "")
            throw Error("File content is empty")
        
        ; Parse Gaussian file content and create ORCA input
        gjfData := ParseGaussianFile(fileContent)
        maxcore := Floor(gjfData["mem"] / gjfData["nprocs"])
        inpFile := RegExReplace(savedPath, "\.gjf$", ".inp")
        
        if (CreateOrcaInput(inpFile, gjfData, maxcore))
        {
            try FileDelete(savedPath)
        }
    }
    catch as e
    {
        MsgBox("Error processing file: " e.Message "`nFile path: " savedPath, "Error", "Icon!")
    }
}
#HotIf

; Function: Wait for save dialog and return saved file path
WaitForSaveDialog()
{
    startTime := A_TickCount
    timeout := 10000 ; 10 second timeout
    
    ; Wait for save dialog to appear
    if (!WinWait("Save As ahk_class #32770",, 5))
        return ""
        
    ; Get the edit control handle
    try {
        saveDialog := WinActive("A")
        editHwnd := ControlGetHwnd("Edit1", saveDialog)
    } catch {
        return ""
    }
    
    ; Wait for dialog to close
    WinWaitClose("Save As ahk_class #32770",, timeout)
    Sleep(200)  ; Give time for file to be written
    
    ; Get the file path that was in the edit control
    try {
        savedFile := ControlGetText("Edit1", "ahk_id " saveDialog)
        if (savedFile && FileExist(savedFile))
            return savedFile
        
        ; If direct path didn't work, try to find file in temp directory
        if (RegExMatch(savedFile, "[^\\]+\.gjf$", &match)) {
            fileName := match[0]
            Loop Files, A_Temp "\gv*\" fileName, "FR"
            {
                if (A_LoopFileTimeModified >= Floor((startTime-2000)/1000))  ; Check if file is new (within last 2 seconds)
                    return A_LoopFileFullPath
            }
            
            ; Also check in user's temp directory
            userTemp := EnvGet("TEMP")
            Loop Files, userTemp "\gv*\" fileName, "FR"
            {
                if (A_LoopFileTimeModified >= Floor((startTime-2000)/1000))
                    return A_LoopFileFullPath
            }
        }
    }
    
    return ""
}

; Function: Parse Gaussian .gjf file
ParseGaussianFile(content)
{
    result := Map()
    
    ; Extract memory
    if (RegExMatch(content, "i)%mem=(\d+)\s*(GB|MB)", &memMatch))
    {
        memValue := Number(memMatch[1])
        memUnit := memMatch[2]
        result["mem"] := memValue * (memUnit = "GB" ? 1000 : 1)
    }
    else
        result["mem"] := 8000 ; Default 8000 MB
    
    ; Extract nprocs
    if (RegExMatch(content, "i)%nprocshared=(\d+)", &procMatch))
        result["nprocs"] := Number(procMatch[1])
    else
        result["nprocs"] := 8 ; Default 8 cores
    
    ; Extract keywords
    if (RegExMatch(content, "m)^#\s+(.+)$", &keyMatch))
    {
        keywords := Trim(keyMatch[1])
        ; Don't remove the ? character, just store it as is
        result["keywords"] := keywords
    }
    
    ; Extract charge and multiplicity
    chargeMultiPattern := "m)^\s*(-?\d+)\s+(\d+)\s*$"
    lines := StrSplit(content, "`n", "`r")
    for line in lines
    {
        if (RegExMatch(line, chargeMultiPattern, &cmMatch))
        {
            result["charge"] := cmMatch[1]
            result["multiplicity"] := cmMatch[2]
            break
        }
    }
    
    ; Extract coordinates
    coords := ""
    inCoords := false
    for line in lines
    {
        if (RegExMatch(line, "^(-?\d+\s+\d+)\s*$"))
        {
            inCoords := true
            continue
        }
        if (inCoords)
        {
            if (Trim(line) = "" || RegExMatch(line, "^%"))
                break
            if (RegExMatch(line, "^\s*[A-Za-z]"))
                coords .= line . "`n"
        }
    }
    result["coordinates"] := RTrim(coords, "`n")
    
    ; Extract other settings
    settings := ""
    inSettings := false
    nestLevel := 0
    
    for line in lines
    {
        if (RegExMatch(line, "^%"))
        {
            if (!RegExMatch(line, "i)^%(mem|nprocshared|chk)="))
            {
                inSettings := true
                settings .= line . "`n"
                nestLevel += StrCount(line, "%")
            }
        }
        else if (inSettings)
        {
            settings .= line . "`n"
            if (InStr(line, "end"))
            {
                nestLevel -= 1
                if (nestLevel = 0)
                    inSettings := false
            }
        }
    }
    result["settings"] := RTrim(settings, "`n")
    
    return result
}

; Function: Create ORCA input file
CreateOrcaInput(filePath, gjfData, maxcore)
{
    content := "%pal nprocs " . gjfData["nprocs"] . " end`n"
    content .= "%maxcore " . maxcore . "`n"
    
    ; Use selected keywords if not "No change"
    keywords := CURRENT_KEYWORDS = "No change" ? gjfData["keywords"] : CURRENT_KEYWORDS
    
    ; Replace ? with / in keywords for ORCA format
    keywords := StrReplace(keywords, "?", "/")
    
    ; Also convert smd=xxx to smd(xxx) format
    keywords := RegExReplace(keywords, "i)\bsmd=([^\s]+)", "smd($1)")
    
    content .= "! " . keywords . "`n"
    
    content .= "*xyz " . gjfData["charge"] . " " . gjfData["multiplicity"] . "`n"
    content .= gjfData["coordinates"] . "`n"
    content .= "*`n"
    
    if (gjfData.Has("settings") && gjfData["settings"] != "")
        content .= gjfData["settings"] . "`n"
    
    try
    {
        if (FileExist(filePath))
            FileDelete(filePath)
        FileAppend(content, filePath)
        return true
    }
    catch as e
    {
        MsgBox("Cannot create ORCA input file: " . e.Message, "Error", "Icon!")
        return false
    }
}

; Helper: Count string occurrences
StrCount(haystack, needle)
{
    count := 0
    pos := 1
    while (pos := InStr(haystack, needle, , pos))
    {
        count++
        pos++
    }
    return count
}

; Function: Get currently selected files in Windows Explorer (modified to return an array)
GetSelectedFiles()
{
    selectedFiles := []
    
    ; Try to get selected files using Windows Explorer
    explorerHwnd := WinExist("ahk_class CabinetWClass") or WinExist("ahk_class ExploreWClass")
    if (explorerHwnd)
    {
        for window in ComObject("Shell.Application").Windows
        {
            try
            {
                if (window.HWND = explorerHwnd)
                {
                    selectedItems := window.Document.SelectedItems
                    for item in selectedItems
                        selectedFiles.Push(item.Path)
                    return selectedFiles
                }
            }
        }
    }
    
    ; If no file is selected in Explorer, return an empty array
    return selectedFiles
}

; Original GetSelectedFile function renamed to GetSelectedFile for backward compatibility if needed
GetSelectedFile()
{
    files := GetSelectedFiles()
    return files.Length > 0 ? files[1] : ""
}

; Function: Process ORCA output file
ProcessOrcaOutputFile(orcaFile, filePath, fileName)
{
    ; Define the Gaussian format log file to be created
    fakeLogFile := filePath . "\" . fileName . "_fake.out"
    
    ; Read ORCA output file content
    fileContent := FileRead(orcaFile)
    
    ; Extract key information
    keywords := ExtractKeywords(fileContent)
    nprocs := ExtractNprocs(fileContent)
    maxcore := ExtractMaxcore(fileContent)
    settings := ExtractSettings(fileContent, "out")
    charge := ExtractCharge(fileContent)
    multiplicity := ExtractMultiplicity(fileContent)
    
    ; Extract absorption spectrum data
    spectrumData := ExtractAbsorptionSpectrum(fileContent)
    
    ; Call OfakeG.exe to convert file
    RunWait("`"" . OFAKE_G_PATH . "`" `"" . orcaFile . "`"", filePath)
    
    ; If generated file exists, add additional information
    if (FileExist(fakeLogFile))
    {
        ; Create content to be added at the beginning of the file
        memGB := Floor(maxcore * nprocs / 1000)
        headerContent := " Entering Link 1 = Welcome! `n"
        headerContent .= "`n%mem=" . memGB . "GB"
        headerContent .= "`n%nprocshared=" . nprocs
        headerContent .= "`n----------------------------------------------------------------------"
        headerContent .= "`n# " . keywords
        headerContent .= "`n----------------------------------------------------------------------"
        headerContent .= "`nUsing 2006 physical constants."
        headerContent .= "`n-------------------"
        headerContent .= "`nTitle Card Required"
        headerContent .= "`n-------------------"
        headerContent .= "`nSymbolic Z-matrix:"
        headerContent .= "`nCharge = " . charge . " Multiplicity = " . multiplicity
        headerContent .= "`n Mg                   -0.00161   2.07295   0.62816"
        headerContent .= "`n"
        
        ; Add other settings to headerContent (possibly for Add. Inp. section)
        if (settings != "")
            headerContent .= "`n" . settings
        
        ; Read existing _fake.out file content
        existingContent := FileRead(fakeLogFile)
        
        ; Format excitation data if found
        footerContent := ""
        if (spectrumData["found"]) {
            footerContent := "`n" . FormatGaussianExcitations(spectrumData)
        }
        
        ; Write new content to file (prepend our content and append excitation data)
        try
        {
            FileDelete(fakeLogFile)
            FileAppend(headerContent . "`n" . existingContent . footerContent, fakeLogFile)
        }
        catch as e
        {
            MsgBox("Cannot modify file: " . e.Message, "Error", "Icon!")
            return
        }
        
        ; Open file with GaussView
        OpenWithGaussView(fakeLogFile)
        
        ; Delete temporary file (if permissions allow)
        try
        {
            Sleep(500)  ; Give GaussView time to open the file
            FileDelete(fakeLogFile)
        }
    }
    else
    {
        MsgBox(fakeLogFile)
        MsgBox("OfakeG.exe failed to generate converted file.", "Error", "Icon!")
    }
}

; Function: Extract keywords from ORCA input and replace / with ?
ExtractKeywords(content)
{
    ; Match pattern like "|  3> ! opt freq wb97x-d3 def2-sv(p) def2-svp/c rijcosx"
    regexPattern := "m)^\s*\|?\s*\d?>?\s*! ?(.+)$"
    if RegExMatch(content, regexPattern, &match)
    {
        result := Trim(match[1])
        result := StrReplace(result, "/", "?") ; Replace / with ?
        return result
    }
    return ""
}

; Function: Extract nprocs value
ExtractNprocs(content)
{
    ; Find nprocs in ORCA settings
    regexPattern := "i)%[ \t]*pal[ \t\r\n]+nprocs[ \t]+(\d+)"
    if (RegExMatch(content, regexPattern, &match))
        return match[1]
    return 1  ; Default value
}

; Function: Extract maxcore value
ExtractMaxcore(content)
{
    ; Find maxcore in ORCA settings
    regexPattern := "i)%[ \t]*maxcore[ \t]+(\d+)"
    if (RegExMatch(content, regexPattern, &match))
        return match[1]
    return 1000  ; Default value (MB)
}

; Function: Extract charge
ExtractCharge(content)
{
    ; Find pattern like *xyz 0 1 where 0 is the charge
    regexPattern := "i)\*xyz[ \t]+(-?\d+)[ \t]+\d+"
    if (RegExMatch(content, regexPattern, &match))
        return match[1]
    return 0  ; Default value
}

; Function: Extract spin multiplicity
ExtractMultiplicity(content)
{
    ; Find pattern like *xyz 0 1 where 1 is the spin multiplicity
    regexPattern := "i)\*xyz[ \t]+(-?\d+)[ \t]+(\d+)"
    if (RegExMatch(content, regexPattern, &match))
        return match[2]
    return 1  ; Default value
}

; Function: Extract molecular coordinates
ExtractCoordinates(content)
{
    coordinates := ""
    
    ; Find coordinates between *xyz and * with improved regex to handle spaces and newlines
    regexPattern := "i)\*\s*xyz[^\n]*\n([\s\S]*?)\*"
    if (RegExMatch(content, regexPattern, &match))
    {
        ; Extract coordinates and trim leading/trailing whitespace
        coordinates := Trim(match[1], " `t`r`n")
    }
    
    return coordinates
}

; Function: Extract other calculation settings
ExtractSettings(content, fileType := "out")
{
    settings := ""
    
    if (fileType = "inp")
    {
        ; Logic for .inp files - handle both single-line and multi-line blocks
        lines := StrSplit(content, "`n", "`r")
        i := 1
        while (i <= lines.Length)
        {
            line := Trim(lines[i])
            
            ; Check if line starts with % and is a setting block
            if (RegExMatch(line, "^%(\w+)(.*)$", &match))
            {
                settingName := match[1]
                restOfLine := Trim(match[2])
                
                ; Skip pal and maxcore as they're handled separately
                if (settingName = "pal" || settingName = "maxcore")
                {
                    i++
                    continue
                }
                
                ; Start building the setting block
                settingBlock := line . "`n"
                
                ; Check if it's a single-line setting ending with "end"
                if (InStr(restOfLine, "end"))
                {
                    settings .= settingBlock
                    i++
                    continue
                }
                
                ; It's a multi-line block, read until we find "end"
                i++
                while (i <= lines.Length)
                {
                    currentLine := lines[i]
                    settingBlock .= currentLine . "`n"
                    
                    ; Check if this line contains "end" (allowing for whitespace)
                    if (RegExMatch(Trim(currentLine), "^end\s*$"))
                        break
                    
                    i++
                }
                
                settings .= settingBlock
            }
            
            i++
        }
    }
    else
    {
        ; Logic for .out files - handle both single-line and multi-line blocks
        lines := StrSplit(content, "`n", "`r")
        i := 1
        while (i <= lines.Length)
        {
            line := lines[i]
            
            ; Check if line matches the .out file format pattern
            if (RegExMatch(line, "^\s*\|\s*\d+>\s*%(\w+)(.*)$", &match))
            {
                settingName := match[1]
                restOfLine := Trim(match[2])
                
                ; Skip pal and maxcore as they're handled separately
                if (settingName = "pal" || settingName = "maxcore")
                {
                    i++
                    continue
                }
                
                ; Start building the setting block
                settingBlock := line . "`n"
                
                ; Check if it's a single-line setting ending with "end"
                if (InStr(restOfLine, "end"))
                {
                    settings .= settingBlock
                    i++
                    continue
                }
                
                ; It's a multi-line block, read until we find "end"
                i++
                while (i <= lines.Length)
                {
                    currentLine := lines[i]
                    settingBlock .= currentLine . "`n"
                    
                    ; Check if this line contains "end" with the .out format
                    if (RegExMatch(currentLine, "^\s*\|\s*\d+>\s*end\s*$"))
                        break
                    
                    i++
                }
                
                settings .= settingBlock
            }
            
            i++
        }
    }
    
    return RTrim(settings, "`n")
}

; Function: Create Gaussian input file
CreateGaussianInput(filePath, nprocs, maxcore, keywords, charge, multiplicity, coordinates, settings)
{
    ; Calculate memory (GB)
    memGB := Floor(maxcore * nprocs / 1000)
    
    ; Build file content
    content := "%mem=" . memGB . "GB`n"
    content .= "%nprocshared=" . nprocs . "`n"
    content .= "# " . keywords . "`n`n"
    content .= "TC`n`n"
    content .= charge . " " . multiplicity . "`n"
    content .= coordinates . "`n`n"
    
    ; Add other settings (if any)
    if (settings != "")
        content .= settings . "`n"
    
    ; Write to file
    try
    {
        ; Get file directory
        SplitPath(filePath, , &fileDir)
        
        ; Create directory if it doesn't exist
        if (!DirExist(fileDir))
            DirCreate(fileDir)
        
        ; Delete file if it exists
        if (FileExist(filePath))
            FileDelete(filePath)
            
        ; Write new file
        FileAppend(content, filePath)
        return true
    }
    catch as e
    {
        MsgBox("Cannot create file: " . filePath . "`nError: " . e.Message, "Error", "Icon!")
        return false
    }
}

; Function: Open file with GaussView
OpenWithGaussView(filePath)
{
    if (FileExist(GAUSS_VIEW_PATH))
    {
        ; Use configured GaussView path to open file
        Run("`"" . GAUSS_VIEW_PATH . "`" `"" . filePath . "`"")
    }
    else
    {
        ; If GaussView is not found, try to open with file association
        Run(filePath)
    }
}

; Function: Extract absorption spectrum data from ORCA output
ExtractAbsorptionSpectrum(content) {
    result := Map()
    result["found"] := false
    
    ; Regular expression to find the absorption spectrum table
    regexPattern := "s)ABSORPTION SPECTRUM VIA TRANSITION ELECTRIC DIPOLE MOMENTS[\s\S]*?-{20,}[\s\S]*?-{20,}([\s\S]*?)-{20,}"
    if (RegExMatch(content, regexPattern, &match)) {
        tableContent := match[1]
        
        ; Now parse the transitions
        transitions := []
        
        ; Split into lines and process each line
        lines := StrSplit(tableContent, "`n", "`r")
        for line in lines {
            ; Skip empty lines
            if (Trim(line) = "") {
                continue
            }
            
            ; Extract transition data with regex
            ; Format: "  0-1A  ->  1-1A    2.849027   22979.0   435.2   1.799267602  25.77751  -5.06692   0.00145  -0.32221"
            transitionPattern := "^\s*\d+-(\d+)(\w+)\s+->\s+(\d+)-(\d+)(\w+)\s+(\d+\.\d+)\s+(\d+\.\d+)\s+(\d+\.\d+)\s+(\d+\.\d+)"
            if (RegExMatch(line, transitionPattern, &tMatch)) {
                transition := Map()
                transition["fromSpin"] := tMatch[1]
                transition["fromSymm"] := tMatch[2]
                transition["state"] := tMatch[3]
                transition["toSpin"] := tMatch[4]
                transition["toSymm"] := tMatch[5]
                transition["energy_eV"] := tMatch[6]
                transition["energy_cm"] := tMatch[7]
                transition["wavelength_nm"] := tMatch[8]
                transition["oscillator"] := tMatch[9]
                
                transitions.Push(transition)
            }
        }
        
        if (transitions.Length > 0) {
            result["found"] := true
            result["transitions"] := transitions
        }
    }
    
    ; Extract S**2 values from STATE lines
    statePattern := "STATE (\d+): E=.*?<S\*\*2> = (\d+\.\d+) Mult (\d+)"
    pos := 1
    while (pos := RegExMatch(content, statePattern, &stateMatch, pos)) {
        stateNum := stateMatch[1]
        s2Value := stateMatch[2]
        multValue := stateMatch[3]
        
        ; Store S**2 values for each state
        if (!result.Has("s2values"))
            result["s2values"] := Map()
            
        result["s2values"][stateNum] := {s2: s2Value, mult: multValue}
        
        pos += stateMatch.Len
    }
    
    return result
}

; Function: Format absorption spectrum data as Gaussian output
FormatGaussianExcitations(spectrumData) {
    if (!spectrumData["found"]) {
        return ""
    }
    
    result := " Excitation energies and oscillator strengths:`n"
    
    transitions := spectrumData["transitions"]
    s2values := spectrumData.Has("s2values") ? spectrumData["s2values"] : Map()
    
    ; Process each transition
    for i, transition in transitions {
        stateNum := transition["state"]
        
        ; Determine multiplicity
        multLabel := "Singlet"
        s2Value := "0.000"
        
        if (s2values.Has(stateNum)) {
            s2Value := s2values[stateNum]["s2"]
            multValue := s2values[stateNum]["mult"]
            
            if (multValue = "3")
                multLabel := "Triplet"
            else if (multValue = "5")
                multLabel := "Quintet"
        }
        
        ; Get values
        energy_eV := transition["energy_eV"]
        wavelength := transition["wavelength_nm"]
        oscillator := transition["oscillator"]
        
        ; Format the line
        result .= Format(" Excited State   {1}:      {2}-{3}      {4} eV  {5} nm  f={6}  <S**2>={7}`n", 
                         i, multLabel, transition["toSymm"], energy_eV, wavelength, oscillator, s2Value)
    }
    
    ; Add the SavETr line
    result .= Format(" SavETr:  write IOETrn=   770 NScale= 10 NData=  16 NLR=1 NState=    {1} LETran=     100.", transitions.Length)
    
    return result
}