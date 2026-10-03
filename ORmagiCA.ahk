#Requires AutoHotkey v2.0

; Global configuration
global OFAKE_G_PATH := "D:\CompChem\OfakeG.exe"
global GAUSS_VIEW_PATH := "D:\GaussView\GV6.0.16WIN\g16w\gview.exe"
global ORCA2GAUSSIAN_PATH := "D:\Projects\Softs\ORmagiCA\ORmagiCA\orca2gaussian.ahk"
global USE_OFAKEG := "1"
global UI_LANG := "zh"          ; UI 语言: zh / en
global SETTINGS_GUI := ""
global CURRENT_KEYWORDS := "No change"

; Initialize on startup
InitializeSettings()

; 按当前 UI 语言返回字符串 (zh=中文, en=English)
L(zh, en) {
    global UI_LANG
    return (UI_LANG = "en") ? en : zh
}

; Initialize settings from INI file
InitializeSettings() {
    global OFAKE_G_PATH, GAUSS_VIEW_PATH, ORCA2GAUSSIAN_PATH, USE_OFAKEG, UI_LANG, CURRENT_KEYWORDS, SETTINGS_GUI
    
    iniPath := A_ScriptDir "\ORmagiCA_settings.ini"
    
    ; Read paths from ini if exists, otherwise keep defaults
    if FileExist(iniPath) {
        savedGV := IniRead(iniPath, "Paths", "GaussView", "")
        savedOF := IniRead(iniPath, "Paths", "OfakeG", "")
        savedO2G := IniRead(iniPath, "Paths", "Orca2Gaussian", "")
        savedUse := IniRead(iniPath, "Paths", "UseOfakeG", "")
        savedLang := IniRead(iniPath, "General", "Language", "")
        
        if (savedGV != "")
            GAUSS_VIEW_PATH := savedGV
        if (savedOF != "")
            OFAKE_G_PATH := savedOF
        if (savedO2G != "")
            ORCA2GAUSSIAN_PATH := savedO2G
        if (savedUse != "")
            USE_OFAKEG := savedUse
        if (savedLang != "")
            UI_LANG := savedLang
    }
    
    ; Create settings GUI
    if !IsObject(SETTINGS_GUI)
        SETTINGS_GUI := SettingsGui()
}

; 从候选列表里选择一个系统中真实存在的字体 (保证兼容性/优雅回退)
PickFont(candidates) {
    for f in candidates
        if (FontExists(f))
            return f
    return candidates[candidates.Length]
}

FontExists(fontName) {
    ; 通过枚举字体键判断该字体是否已安装
    keys := "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
    Loop Reg, keys, "V"
        if (InStr(A_LoopRegName, fontName, true) != 0)
            return true
    return false
}

; Settings GUI Class
class SettingsGui {
    __New() {
        ; Create window with shadow style (CS_DROPSHADOW = 0x20000)
        this.gui := Gui("+AlwaysOnTop -Caption +ToolWindow +LastFound")
        DllCall("SetClassLong", "Ptr", this.gui.Hwnd, "Int", -26, "Int", DllCall("GetClassLong", "Ptr", this.gui.Hwnd, "Int", -26) | 0x20000)

        ; 优雅古典衬线字体 (带兼容回退)
        this.fontName := PickFont(["Palatino Linotype", "Book Antiqua", "Georgia", "Cambria", "Times New Roman"])
        this.gui.SetFont("s11", this.fontName)
        this.gui.BackColor := "F4F1EC"   ; 米白古典底

        ; 顶部标题条
        this.lblTitle := this.gui.Add("Text", "x14 y12 w400 Center", L("⚙  ORmagiCA 设置", "⚙  ORmagiCA  Settings"))
        this.gui.SetFont("s12 Bold", this.fontName)
        this.lblTitle.SetFont("s12 Bold", this.fontName)
        this.gui.SetFont("s11", this.fontName)
        this.lblSep := this.gui.Add("Text", "x14 y42 w400 Center", "──────────────────────────────")

        ; 标签
        this.gui.SetFont("s10", this.fontName)
        this.lblGv := this.gui.Add("Text", "x14 y62", L("GaussView 路径:", "GaussView Path:"))
        this.gvPath := this.gui.Add("Edit", "x14 y82 w400", GAUSS_VIEW_PATH)

        this.lblOf := this.gui.Add("Text", "x14 y112", L("OfakeG 路径:", "OfakeG Path:"))
        this.ofakePath := this.gui.Add("Edit", "x14 y132 w400", OFAKE_G_PATH)

        this.lblO2g := this.gui.Add("Text", "x14 y162", L("orca2gaussian 模块路径 (.ahk / .py):", "orca2gaussian module path (.ahk / .py):"))
        this.o2gPath := this.gui.Add("Edit", "x14 y182 w400", ORCA2GAUSSIAN_PATH)

        this.lblUse := this.gui.Add("Text", "x14 y212", L("Use OfakeG (1=OfakeG, 0=orca2gaussian):", "Use OfakeG (1=OfakeG, 0=orca2gaussian):"))
        this.useOfakeG := this.gui.Add("Edit", "x14 y232 w400", USE_OFAKEG)

        this.lblLang := this.gui.Add("Text", "x14 y262", L("UI 语言 (zh/en):", "UI Language (zh/en):"))
        this.langDrop := this.gui.Add("DropDownList", "x14 y282 w400", ["zh", "en"])
        this.langDrop.Value := (UI_LANG = "en") ? 2 : 1

        ; ORCA Keywords section with "No change" as protected item
        this.lblKw := this.gui.Add("Text", "x14 y312", L("ORCA 关键字:", "ORCA Keywords:"))
        this.keywordsList := this.gui.Add("ListBox", "x14 y332 w400 h150 vKeywordsList")

        ; Add keyword input and buttons (衬线字体按钮)
        this.gui.SetFont("s10", this.fontName)
        this.newKeyword := this.gui.Add("Edit", "x14 y502 w300")
        this.btnAdd := this.gui.Add("Button", "x320 y502 w94 h28", L("添加", "Add"))
        this.btnDel := this.gui.Add("Button", "x320 y536 w94 h28", L("删除", "Delete"))

        ; Events
        this.gvPath.OnEvent("Change", this.SavePaths.Bind(this))
        this.ofakePath.OnEvent("Change", this.SavePaths.Bind(this))
        this.o2gPath.OnEvent("Change", this.SavePaths.Bind(this))
        this.useOfakeG.OnEvent("Change", this.SavePaths.Bind(this))
        this.langDrop.OnEvent("Change", this.OnLangChange.Bind(this))
        this.keywordsList.OnEvent("Change", this.UpdateKeywords.Bind(this))
        this.keywordsList.OnEvent("Change", this.OnKeywordSelected.Bind(this))  ; Add this line
        this.btnAdd.OnEvent("Click", this.AddKeyword.Bind(this))
        this.btnDel.OnEvent("Click", this.DeleteKeyword.Bind(this))
        
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
        winWidth := 436
        winHeight := 580
        
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
        global OFAKE_G_PATH, GAUSS_VIEW_PATH, ORCA2GAUSSIAN_PATH, USE_OFAKEG
        
        iniPath := A_ScriptDir "\ORmagiCA_settings.ini"
        
        GAUSS_VIEW_PATH := this.gvPath.Text
        OFAKE_G_PATH := this.ofakePath.Text
        ORCA2GAUSSIAN_PATH := this.o2gPath.Text
        USE_OFAKEG := this.useOfakeG.Text
        
        IniWrite(GAUSS_VIEW_PATH, iniPath, "Paths", "GaussView")
        IniWrite(OFAKE_G_PATH, iniPath, "Paths", "OfakeG")
        IniWrite(ORCA2GAUSSIAN_PATH, iniPath, "Paths", "Orca2Gaussian")
        IniWrite(USE_OFAKEG, iniPath, "Paths", "UseOfakeG")
    }
    
    OnLangChange(*) {
        global UI_LANG
        UI_LANG := (this.langDrop.Value = 2) ? "en" : "zh"
        iniPath := A_ScriptDir "\ORmagiCA_settings.ini"
        IniWrite(UI_LANG, iniPath, "General", "Language")
        ; 就地更新标签文本，避免销毁重建冲突
        try {
            this.lblTitle.Text := L("⚙  ORmagiCA 设置", "⚙  ORmagiCA  Settings")
            this.lblGv.Text := L("GaussView 路径:", "GaussView Path:")
            this.lblOf.Text := L("OfakeG 路径:", "OfakeG Path:")
            this.lblO2g.Text := L("orca2gaussian 模块路径 (.ahk / .py):", "orca2gaussian module path (.ahk / .py):")
            this.lblUse.Text := L("Use OfakeG (1=OfakeG, 0=orca2gaussian):", "Use OfakeG (1=OfakeG, 0=orca2gaussian):")
            this.lblLang.Text := L("UI 语言 (zh/en):", "UI Language (zh/en):")
            this.lblKw.Text := L("ORCA 关键字:", "ORCA Keywords:")
            this.btnAdd.Text := L("添加", "Add")
            this.btnDel.Text := L("删除", "Delete")
        }
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
    else
        result["keywords"] := ""
    
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
    
    ; Extract coordinates (with optional freeze/fragment flags) and the
    ; opt=modredundant section that follows them
    atoms := []
    frozen := []          ; 0-based indices of atoms frozen via -1 flag
    cleanCoords := ""
    inCoords := false
    coordDone := false
    modredLines := []
    for line in lines
    {
        if (!inCoords && RegExMatch(line, "^\s*-?\d+\s+\d+\s*$"))
        {
            inCoords := true
            continue
        }
        if (inCoords && !coordDone)
        {
            trimmedLine := Trim(line)
            if (trimmedLine = "" || RegExMatch(line, "^%"))
            {
                coordDone := true
                continue
            }
            atom := ParseAtomLine(trimmedLine)
            if (IsObject(atom))
            {
                atoms.Push(atom)
                if (atom["frozen"])
                    frozen.Push(atoms.Length - 1)
                cleanCoords .= atom["sym"] . " " . atom["x"] . " " . atom["y"] . " " . atom["z"] . "`n"
            }
        }
        else if (coordDone)
        {
            trimmedLine := Trim(line)
            if (trimmedLine = "")
            {
                if (modredLines.Length > 0)
                    break   ; blank line after the modredundant section -> done
                continue
            }
            if (IsModRedundantCandidate(trimmedLine))
                modredLines.Push(trimmedLine)
            else if (modredLines.Length > 0)
                break   ; some other section started
        }
    }
    result["coordinates"] := RTrim(cleanCoords, "`n")
    result["atoms"] := atoms
    result["frozen"] := frozen
    
    ; Parse the modredundant section into constraints and scans
    constraints := []
    scans := []
    for modredLine in modredLines
    {
        entry := ParseModRedundantLine(modredLine)
        if (IsObject(entry))
        {
            if (entry["kind"] = "scan")
                scans.Push(entry)
            else
                constraints.Push(entry)
        }
    }
    result["modredConstraints"] := constraints
    result["modredScans"] := scans
    result["hasModRedKeyword"] := RegExMatch(result["keywords"], "i)modredundant") ? true : false
    
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
    
    ; opt=modredundant is Gaussian-only; in ORCA the constraints/scans
    ; live in the %geom block instead
    if (gjfData.Has("hasModRedKeyword") && gjfData["hasModRedKeyword"])
        keywords := CleanupModRedundantKeyword(keywords)
    
    content .= "! " . keywords . "`n"

    ; 把 % 段落关键字 (如 %geom ... end) 放在 "!" 关键字行与几何坐标之间，便于查看
    if (gjfData.Has("settings") && gjfData["settings"] != "")
        content .= "`n" . gjfData["settings"] . "`n"

    ; Convert G16 modredundant constraints/scans and frozen atoms to ORCA format
    geomBlock := BuildOrcaGeomBlock(gjfData)
    if (geomBlock != "")
        content .= "`n" . geomBlock . "`n"

    content .= "*xyz " . gjfData["charge"] . " " . gjfData["multiplicity"] . "`n"
    content .= gjfData["coordinates"] . "`n"
    content .= "*`n"

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
    
    ; Method 1: Get selected files via ShellWindows collection
    explorerHwnd := WinExist("ahk_class CabinetWClass") or WinExist("ahk_class ExploreWClass")
    if (explorerHwnd)
    {
        try
        {
            for window in ComObject("Shell.Application").Windows
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
    
    ; Method 2 (fallback): If ShellWindows is broken (e.g. after Explorer restart it can
    ; enumerate zero windows), get the selection from the active Explorer window by copying
    ; it to the clipboard and reading the CF_HDROP data. This does not depend on ShellWindows.
    if (selectedFiles.Length = 0 && WinActive("ahk_class CabinetWClass"))
    {
        selectedFiles := GetSelectedFilesViaClipboard()
        if (selectedFiles.Length > 0)
            return selectedFiles
    }
    
    ; If no file is selected in Explorer, return an empty array
    return selectedFiles
}

; Fallback: get selected files by copying the Explorer selection to the clipboard
; and reading the CF_HDROP data (works even when ShellWindows returns no windows)
GetSelectedFilesViaClipboard()
{
    files := []
    
    ; Only run when an Explorer window is active to avoid touching other apps
    if !WinActive("ahk_class CabinetWClass")
        return files
    
    ; Copy the current selection
    Send("^c")
    
    loop 5
    {
        Sleep(100)
        if DllCall("OpenClipboard", "Ptr", 0)
        {
            hDrop := DllCall("GetClipboardData", "UInt", 15, "Ptr")  ; CF_HDROP = 15
            if (hDrop)
            {
                fileCount := DllCall("shell32\DragQueryFileW", "Ptr", hDrop, "UInt", 0xFFFFFFFF, "Ptr", 0, "UInt", 0)
                Loop fileCount
                {
                    charCount := DllCall("shell32\DragQueryFileW", "Ptr", hDrop, "UInt", A_Index - 1, "Ptr", 0, "UInt", 0)
                    buf := Buffer((charCount + 1) * 2, 0)
                    DllCall("shell32\DragQueryFileW", "Ptr", hDrop, "UInt", A_Index - 1, "Ptr", buf, "UInt", charCount + 1)
                    files.Push(StrGet(buf, charCount, "UTF-16"))
                }
            }
            DllCall("CloseClipboard")
            if (files.Length > 0)
                return files
        }
    }
    
    return files
}

; Original GetSelectedFile function renamed to GetSelectedFile for backward compatibility if needed
GetSelectedFile()
{
    files := GetSelectedFiles()
    return files.Length > 0 ? files[1] : ""
}

; ---------------------------------------------------------------------
; RunWithProgress: 显示不确定进度条 (marquee)，运行命令并等待其结束。
;   用 Run 启动并轮询进程 PID，同时推进进度条动画，转换完成后关闭。
; ---------------------------------------------------------------------
RunWithProgress(cmd, workDir, titleText)
{
    progGui := Gui("+AlwaysOnTop +ToolWindow -SysMenu")
    progGui.SetFont("s9", "Segoe UI")
    progGui.SetFont("s10 Bold", "Segoe UI")
    progGui.Add("Text", "x14 y14 w360 Center", titleText)
    progGui.SetFont("s9", "Segoe UI")
    progGui.Add("Text", "x14 y44 w360 Center vStatusText", L("正在处理...", "Processing..."))
    bar := progGui.Add("Progress", "x14 y74 w360 h26 Range0-100")
    progGui.Show("w388 h120")

    ; 启动转换 (非阻塞)
    started := A_TickCount
    Run(cmd, workDir, , &procPID)

    ; 推进 marquee 动画
    p := 0
    While ProcessExist(procPID)
    {
        p += 4
        if (p > 100)
            p := 0
        try bar.Value := p
        Sleep(40)
    }
    ; 至少显示一小段时间，避免一闪而过
    elapsed := A_TickCount - started
    while (elapsed < 600) {
        Sleep(30)
        elapsed := A_TickCount - started
    }
    if IsSet(procPID)
        p := 0
    try progGui.Destroy()
}

; Function: Process ORCA output file
ProcessOrcaOutputFile(orcaFile, filePath, fileName)
{
    global USE_OFAKEG, ORCA2GAUSSIAN_PATH, OFAKE_G_PATH
    
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
    
    ; Choose converter based on UseOfakeG setting
    if (USE_OFAKEG = "0") {
        ; Use orca2gaussian converter (.ahk or .py)
        ; NOTE: orca2gaussian writes <input>_fake.log. Remove any old _fake.out.
        if FileExist(fakeLogFile)
            FileDelete(fakeLogFile)
        ; If the job is a relaxed surface scan, orca2gaussian shows a popup to choose
        ; scanall / scanlast, so pass nothing extra and let it prompt.
        ; 根据扩展名选择解释器: .ahk 用 A_AhkPath, .py 用 python
        SplitPath(ORCA2GAUSSIAN_PATH, , , &o2gExt)
        if (o2gExt = "py")
            cmd := "python `"" . ORCA2GAUSSIAN_PATH . "`" `"" . orcaFile . "`""
        else
            cmd := "`"" . A_AhkPath . "`" `"" . ORCA2GAUSSIAN_PATH . "`" `"" . orcaFile . "`""
        RunWithProgress(cmd, filePath, L("orca2gaussian 转换中...", "orca2gaussian converting..."))
        generatedLog := RegExReplace(orcaFile, "\.out$") . "_fake.log"
        if FileExist(generatedLog) {
            ; 交给 GaussView 前把 route 行里的 "/" 转义为 "?"，GV 才不至于拆分基组关键字
            ; (例如 def2/j -> def2?j)。原始 orca2gaussian 产物仍保留 "/"。
            EscGvRoute(generatedLog)
            OpenWithGaussView(generatedLog)
            try {
                Sleep(500)
                FileDelete(generatedLog)
            }
            return
        }
        MsgBox("orca2gaussian failed to generate a converted file.", "Error", "Icon!")
        return
    }
    
    ; Call OfakeG.exe to convert file
    RunWithProgress("`"" . OFAKE_G_PATH . "`" `"" . orcaFile . "`"", filePath, L("OfakeG 转换中...", "OfakeG converting..."))
    
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

; Function: Extract keywords from ORCA input (keep slashes verbatim)
ExtractKeywords(content)
{
    ; Match pattern like "|  3> ! opt freq wb97x-d3 def2-sv(p) def2-svp/c rijcosx"
    regexPattern := "m)^\s*\|?\s*\d?>?\s*! ?(.+)$"
    if RegExMatch(content, regexPattern, &match)
    {
        return Trim(match[1])
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

; 统计字符串前导空白字符个数 (用于 ORCA 块嵌套终止判断)
LeadWS(s)
{
    n := 0
    While (n < StrLen(s) && (SubStr(s, n + 1, 1) = " " || SubStr(s, n + 1, 1) = "`t"))
        n++
    return n
}

; 判断 ORCA 段内容是否为开启嵌套 end 的子段 (如 %geom 内的 Constraints/Scan)
IsOrcaSubSection(txt)
{
    t := Trim(txt)
    return RegExMatch(t, "i)^(Constraints|Scan|DIIS|SOSCF|NewGTO|NewAuxGTO|NewECP|ReducedInternalCoordinates)$") ? true : false
}

; Function: Extract other calculation settings
ExtractSettings(content, fileType := "out")
{
    settings := ""
    
    if (fileType = "inp")
    {
        ; Logic for .inp files - handle both single-line and multi-line (nested) blocks
        lines := StrSplit(content, "`n", "`r")
        i := 1
        while (i <= lines.Length)
        {
            rawLine := lines[i]
            line := Trim(rawLine)

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
                settingBlock := rawLine . "`n"

                ; Check if it's a single-line setting ending with "end"
                if (RegExMatch(restOfLine, "i)\bend\b"))
                {
                    settings .= settingBlock
                    i++
                    continue
                }

                ; Multi-line block: collect until the matching "end" (nesting-aware).
                ; ORCA 块可嵌套(如 %geom 内的 Constraints ... end)，不能见到第一个 end 就停。
                baseIndent := LeadWS(rawLine)
                depth := 1
                i++
                while (i <= lines.Length)
                {
                    curRaw := lines[i]
                    cur := Trim(curRaw)
                    ; 安全：遇到新的顶层指令(%/!/几何 *xyz/*)则停止，绝不把几何坐标并入设置
                    if (RegExMatch(cur, "^(%|!|\*)"))
                        break
                    settingBlock .= curRaw . "`n"
                    if (RegExMatch(cur, "i)^end$"))
                    {
                        depth--
                        if (depth <= 0 || LeadWS(curRaw) <= baseIndent)
                        {
                            i++
                            break
                        }
                    }
                    else if (IsOrcaSubSection(cur))
                        depth++
                    i++
                }

                settings .= settingBlock
                continue
            }

            i++
        }
    }
    else
    {
        ; Logic for .out files - handle both single-line and multi-line (nested) blocks
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
                if (RegExMatch(restOfLine, "i)\bend\b"))
                {
                    settings .= settingBlock
                    i++
                    continue
                }

                ; Multi-line block: collect until the matching "end" (nesting-aware).
                ; 回显行的缩进不可靠，故用子段计数判断嵌套深度。
                depth := 1
                i++
                while (i <= lines.Length)
                {
                    curLine := lines[i]
                    curContent := ""
                    if (RegExMatch(curLine, "^\s*\|\s*\d+>\s*(.*)$", &cm))
                        curContent := cm[1]
                    cur := Trim(curContent)
                    ; 安全：遇到新的顶层指令则停止(说明块未正确闭合)
                    if (RegExMatch(cur, "^(%|!|\*)"))
                        break
                    settingBlock .= curLine . "`n"
                    if (RegExMatch(cur, "i)^end$"))
                    {
                        depth--
                        if (depth <= 0)
                        {
                            i++
                            break
                        }
                    }
                    else if (IsOrcaSubSection(cur))
                        depth++
                    i++
                }

                settings .= settingBlock
                continue
            }

            i++
        }
    }
    
    return RTrim(settings, "`n")
}

; 用换行连接数组
JoinLines(arr)
{
    s := ""
    for l in arr
        s .= l . "`n"
    return s
}

; 把 ORCA %geom 块里的 Constraints/Scan 子段转成 GaussView 可读的 G16 ModRedundant 行。
;   { B i j C } / { A i j k C } / { D i j k l C } -> "B i+1 j+1 F" ... (0-based -> 1-based)
;   { C i C }                                     -> "X i+1 F"
;   T a1 [a2 a3 a4] = start, end, n               -> "T a1+1 ... S n step"
; geomLines: %geom 块的所有行(含首行 %geom 与末行 end)
; modredArr: 输出数组, 追加 G16 ModRedundant 行
; 返回: 移除 Constraints/Scan 子段后的 %geom 文本(若已无实质内容则返回 "")
ParseGeomBlockToModred(geomLines, modredArr)
{
    out := []
    n := geomLines.Length
    i := 1
    while (i <= n)
    {
        t := Trim(geomLines[i])
        if (RegExMatch(t, "i)^Constraints$"))
        {
            i++
            while (i <= n)
            {
                ct := Trim(geomLines[i])
                if (RegExMatch(ct, "i)^end$"))
                {
                    i++
                    break
                }
                if (RegExMatch(ct, "^\{\s*([BAD])\s+([\d\s]+?)\s*C\s*\}$", &m))
                {
                    typ := StrUpper(m[1])
                    ml := typ
                    for tok in StrSplit(Trim(m[2]), " ")
                        if (tok != "")
                            ml .= " " . (Integer(tok) + 1)
                    ml .= " F"
                    modredArr.Push(ml)
                }
                else if (RegExMatch(ct, "^\{\s*C\s+(\d+)\s+C\s*\}$", &mc))
                    modredArr.Push("X " . (Integer(mc[1]) + 1) . " F")
                i++
            }
            continue
        }
        else if (RegExMatch(t, "i)^Scan$"))
        {
            i++
            while (i <= n)
            {
                st := Trim(geomLines[i])
                if (RegExMatch(st, "i)^end$"))
                {
                    i++
                    break
                }
                if (RegExMatch(st, "i)^([BAD])\s+([\d\s]+?)\s*=\s*(-?[\d\.]+)\s*,\s*(-?[\d\.]+)\s*,\s*(\d+)\s*$", &ms))
                {
                    typ := StrUpper(ms[1])
                    ml := typ
                    for tok in StrSplit(Trim(ms[2]), " ")
                        if (tok != "")
                            ml .= " " . (Integer(tok) + 1)
                    nstep := Integer(ms[5])
                    step := (ms[4] + 0 - (ms[3] + 0)) / Max(nstep, 1)
                    ml .= " S " . nstep . " " . Format("{:.6f}", step)
                    modredArr.Push(ml)
                }
                i++
            }
            continue
        }
        out.Push(geomLines[i])
        i++
    }
    ; 若除 %geom/end 外无实质内容, 视为整块已转换
    hasContent := false
    for l in out
    {
        tl := Trim(l)
        if (tl != "" && !RegExMatch(tl, "i)^%geom$") && !RegExMatch(tl, "i)^end$"))
            hasContent := true
    }
    return hasContent ? RTrim(JoinLines(out), "`n") : ""
}

; 扫描 settings 文本, 把 %geom 的 Constraints/Scan 转为 G16 ModRedundant 行。
; 返回 Map: { modred: "...", settings: "移除已转换子段后的 settings" }
ConvertOrcaGeomToModRedundant(settings)
{
    modred := []
    remaining := []
    lines := StrSplit(settings, "`n", "`r")
    n := lines.Length
    i := 1
    while (i <= n)
    {
        line := lines[i]
        if (RegExMatch(Trim(line), "i)^%geom\b"))
        {
            geomLines := [line]
            depth := 1
            i++
            while (i <= n)
            {
                cur := lines[i]
                t := Trim(cur)
                if (RegExMatch(t, "^(%|!|\*)"))
                    break
                geomLines.Push(cur)
                if (RegExMatch(t, "i)^end$"))
                {
                    depth--
                    if (depth <= 0)
                    {
                        i++
                        break
                    }
                }
                else if (IsOrcaSubSection(t))
                    depth++
                i++
            }
            keep := ParseGeomBlockToModred(geomLines, modred)
            if (keep != "")
                remaining.Push(keep)
            continue
        }
        remaining.Push(line)
        i++
    }
    res := Map()
    res["modred"] := modred.Length ? RTrim(JoinLines(modred), "`n") : ""
    res["settings"] := RTrim(JoinLines(remaining), "`n")
    return res
}

; Function: Create Gaussian input file
CreateGaussianInput(filePath, nprocs, maxcore, keywords, charge, multiplicity, coordinates, settings)
{
    ; Calculate memory (GB)
    memGB := Floor(maxcore * nprocs / 1000)

    ; 把 ORCA %geom 的 Constraints/Scan 转成 G16 ModRedundant(供 GV 读取冻结/扫描)
    conv := ConvertOrcaGeomToModRedundant(settings)
    modred := conv["modred"]
    settings := conv["settings"]

    ; route: 关键字里的 "/" 转义为 "?" (避免 GV 拆分基组)，保存回 ORCA 时再还原
    routeKw := StrReplace(keywords, "/", "?")
    ; 若有冻结/扫描, 确保 route 含 opt=modredundant, 否则 GV 不显示约束编辑器
    if (modred != "" && !RegExMatch(routeKw, "i)modredundant"))
    {
        parts := StrSplit(routeKw, " ")
        foundOpt := false
        for idx, p in parts
        {
            if (StrLower(p) = "opt")
            {
                parts[idx] := "opt=modredundant"
                foundOpt := true
            }
        }
        if (!foundOpt)
            parts.Push("opt=modredundant")
        routeKw := ""
        for p in parts
            routeKw .= (routeKw = "" ? "" : " ") . p
    }

    ; Build file content
    content := "%mem=" . memGB . "GB`n"
    content .= "%nprocshared=" . nprocs . "`n"
    content .= "# " . routeKw . "`n`n"
    content .= "TC`n`n"
    content .= charge . " " . multiplicity . "`n"
    content .= coordinates . "`n`n"

    ; G16 ModRedundant 行放在坐标之后的空行之后(独立的一段)，GV 才能正确解析
    if (modred != "")
        content .= modred . "`n`n"

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

; 在把 .log 交给 GaussView 前，将 route 行(以 " # " 开头)里辅助基的 "/" 转义为 "?"，
; 避免 GV 拆分基组关键字(如 def2/j 显示成 def2 def2-svp j)。保留第一个 "/"(method/basis 分隔)。
EscGvRoute(logFile)
{
    if !FileExist(logFile)
        return
    lines := StrSplit(FileRead(logFile), "`n", "`r")
    out := []
    For idx, line in lines {
        if (SubStr(line, 1, 3) = " # ") {
            ; 按 "/" 分段：method/basis 间的第一个 "/" 保留，其后辅助基的 "/" 改为 "?"
            segments := StrSplit(line, "/")
            newLine := segments[1] . "/" . segments[2]
            s := 2
            While (++s <= segments.Length)
                newLine .= "?" . segments[s]
            line := newLine
        }
        out.Push(line)
    }
    try {
        txt := ""
        For ln in out
            txt .= ln . "`n"
        FileDelete(logFile)
        FileAppend(txt, logFile)
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

; ---------------------------------------------------------------------------
; Gaussian modredundant -> ORCA %geom conversion
; ---------------------------------------------------------------------------

; Split a line into whitespace-separated tokens
TokenizeLine(line)
{
    tokens := []
    pos := 1
    while (pos := RegExMatch(line, "\S+", &m, pos))
    {
        tokens.Push(m[0])
        pos += m.Len
    }
    return tokens
}

IsNumericToken(tok)
{
    return RegExMatch(tok, "^[-+]?(\d+(\.\d*)?|\.\d+)([EeDd][-+]?\d+)?$") ? true : false
}

; Parse one coordinate line of a .gjf file.
; Supported forms:
;   Sym x y z            (plain atom)
;   Sym flag x y z       (flag: 0 = free, -1 = frozen, >0 = fragment id)
ParseAtomLine(line)
{
    tokens := TokenizeLine(line)
    if (tokens.Length < 4)
        return ""
    
    sym := tokens[1]
    if !RegExMatch(sym, "^[A-Za-z]")
        return ""
    
    ; Coordinates are the last three tokens; anything between the symbol and
    ; them is a fragment number or freeze flag (e.g. 0 / -1)
    xTok := tokens[tokens.Length - 2]
    yTok := tokens[tokens.Length - 1]
    zTok := tokens[tokens.Length]
    if !(IsNumericToken(xTok) && IsNumericToken(yTok) && IsNumericToken(zTok))
        return ""
    
    frozen := false
    if (tokens.Length > 4)
    {
        if !IsNumericToken(tokens[2])
            return ""
        frozen := (Integer(tokens[2]) < 0)
    }
    
    atom := Map()
    atom["sym"] := sym
    atom["x"] := xTok
    atom["y"] := yTok
    atom["z"] := zTok
    atom["frozen"] := frozen
    return atom
}

; Quick check whether a line looks like a modredundant directive,
; e.g. "B 20 7 F", "A 93 91 94 F", "D 22 39 53 91 F", "A 92 91 94 S 10 0.1"
IsModRedundantCandidate(line)
{
    return RegExMatch(line, "i)^[BADX]\s+\d+(\s+\d+)*\s+[FS](\s|$)") ? true : false
}

; Parse a single G16 modredundant directive line.
;   B i j F | A i j k F | D i j k l F | X i F      -> constraint entries
;   B/A/D ... S steps stepSize                      -> scan entry
; Returns "" for unsupported lines (K/break, value-only lines, etc.)
ParseModRedundantLine(line)
{
    tokens := TokenizeLine(line)
    if (tokens.Length < 2)
        return ""
    
    t1 := tokens[1]
    if (t1 = "B" || t1 = "b")
        type := "B", expectedInts := 2
    else if (t1 = "A" || t1 = "a")
        type := "A", expectedInts := 3
    else if (t1 = "D" || t1 = "d")
        type := "D", expectedInts := 4
    else if (t1 = "X" || t1 = "x")
        type := "X", expectedInts := 0
    else
        return ""
    
    ; Cartesian freeze: only "X atom F" is supported -> { C N C }
    if (type = "X")
    {
        if (tokens.Length != 3 || !RegExMatch(tokens[2], "^\d+$") || !(tokens[3] = "F" || tokens[3] = "f"))
            return ""
        entry := Map()
        entry["kind"] := "constraint"
        entry["type"] := "X"
        entry["atoms"] := [Integer(tokens[2])]
        return entry
    }
    
    minLen := 1 + expectedInts + 1   ; type + indices + modifier
    if (tokens.Length < minLen)
        return ""
    
    indices := []
    i := 2
    Loop expectedInts
    {
        if !RegExMatch(tokens[i], "^\d+$")
            return ""
        indices.Push(Integer(tokens[i]))
        i++
    }
    
    modifier := tokens[i]
    entry := Map()
    entry["type"] := type
    entry["atoms"] := indices
    
    if (modifier = "F" || modifier = "f")
    {
        entry["kind"] := "constraint"
    }
    else if (modifier = "S" || modifier = "s")
    {
        if (tokens.Length < i + 2)
            return ""
        stepsTok := tokens[i + 1]
        stepTok := tokens[i + 2]
        if !(RegExMatch(stepsTok, "^\d+$") && IsNumericToken(stepTok))
            return ""
        entry["kind"] := "scan"
        entry["steps"] := Integer(stepsTok)
        entry["stepSize"] := Number(stepTok)
    }
    else
        return ""
    
    return entry
}

; Build an ORCA %geom block from parsed gjf data.
; Returns "" when there is nothing to constrain or scan.
BuildOrcaGeomBlock(gjfData)
{
    constraintLines := []
    scanLines := []
    
    ; Redundant internal-coordinate freezes (B/A/D ... F) and cartesian
    ; freezes (X ... F); G16 numbering is 1-based, ORCA is 0-based
    for entry in gjfData["modredConstraints"]
    {
        if (entry["type"] = "X")
            constraintLines.Push("{ C " . (entry["atoms"][1] - 1) . " C }")
        else
        {
            line := "{ " . entry["type"]
            for v in entry["atoms"]
                line .= " " . (v - 1)
            line .= " C }"
            constraintLines.Push(line)
        }
    }
    
    ; Atoms frozen via the -1 flag in the coordinate block -> { C N C }
    for orcaIdx in gjfData["frozen"]
    {
        line := "{ C " . orcaIdx . " C }"
        if !HasArrayValue(constraintLines, line)
            constraintLines.Push(line)
    }
    
    ; Scans: G16 reads the start value from the input geometry and only gives
    ; steps/step size, so compute start (and end) from the coordinates here
    atoms := gjfData["atoms"]
    for entry in gjfData["modredScans"]
    {
        idx := entry["atoms"]
        switch entry["type"]
        {
            case "B": startVal := GeomBond(atoms, idx[1], idx[2])
            case "A": startVal := GeomAngle(atoms, idx[1], idx[2], idx[3])
            case "D": startVal := GeomDihedral(atoms, idx[1], idx[2], idx[3], idx[4])
            default:  startVal := ""
        }
        
        if !IsNumber(startVal)
            continue   ; invalid atom reference in this scan line
        
        endVal := startVal + entry["steps"] * entry["stepSize"]
        
        line := entry["type"]
        for v in idx
            line .= " " . (v - 1)
        line .= " = " . FormatOrcaValue(startVal) . ", " . FormatOrcaValue(endVal) . ", " . entry["steps"]
        scanLines.Push(line)
    }
    
    if (constraintLines.Length = 0 && scanLines.Length = 0)
        return ""
    
    block := "%geom`n"
    if (constraintLines.Length > 0)
    {
        block .= "      Constraints`n"
        for cl in constraintLines
            block .= "          " . cl . "`n"
        block .= "      end`n"
    }
    if (scanLines.Length > 0)
    {
        block .= "      Scan`n"
        for sl in scanLines
            block .= "          " . sl . "`n"
        block .= "      end`n"
    }
    block .= " end"
    return block
}

HasArrayValue(arr, val)
{
    for v in arr
        if (v = val)
            return true
    return false
}

FormatOrcaValue(v)
{
    return Format("{:.6f}", v)
}

; Remove the Gaussian-only modredundant keyword variants, keeping any other
; opt options intact (e.g. "opt=(modredundant,maxcycle=50)" -> "opt=(maxcycle=50)")
CleanupModRedundantKeyword(keywords)
{
    keywords := RegExReplace(keywords, "i)\bopt\s*=\s*\(\s*modredundant\s*,\s*", "opt=(")
    keywords := RegExReplace(keywords, "i)\bopt\s*=\s*\(\s*modredundant\s*\)", "opt")
    keywords := RegExReplace(keywords, "i)\bopt\s*=\s*modredundant\s*,\s*(?=[^\s(])", "opt=(")
    keywords := RegExReplace(keywords, "i)\bopt\s*=\s*modredundant\b", "opt")
    keywords := RegExReplace(keywords, "i)(^|\s)modredundant(\s|$)", "$1$2")
    keywords := Trim(RegExReplace(keywords, "\s+", " "))
    return keywords
}

; ---------------------------------------------------------------------------
; Geometry helpers (G16 1-based atom indices; vectors are [x, y, z] arrays)
; ---------------------------------------------------------------------------

AtomPos(atoms, g16Idx)
{
    if (g16Idx < 1 || g16Idx > atoms.Length)
        return ""
    a := atoms[g16Idx]
    return [Number(a["x"]), Number(a["y"]), Number(a["z"])]
}

GeomBond(atoms, i, j)
{
    a := AtomPos(atoms, i), b := AtomPos(atoms, j)
    if (!IsObject(a) || !IsObject(b))
        return ""
    d := VecSub(a, b)
    return Sqrt(d[1]*d[1] + d[2]*d[2] + d[3]*d[3])
}

; Angle at central atom j
GeomAngle(atoms, i, j, k)
{
    a := AtomPos(atoms, i), b := AtomPos(atoms, j), c := AtomPos(atoms, k)
    if (!(IsObject(a) && IsObject(b) && IsObject(c)))
        return ""
    u := VecSub(a, b)
    v := VecSub(c, b)
    lu := VecNorm(u), lv := VecNorm(v)
    if (lu = 0 || lv = 0)
        return ""
    cosV := VecDot(u, v) / (lu * lv)
    cosV := Min(1, Max(-1, cosV))
    return RadToDeg(ACos(cosV))
}

; Torsion i-j-k-l
GeomDihedral(atoms, i, j, k, l)
{
    p1 := AtomPos(atoms, i), p2 := AtomPos(atoms, j), p3 := AtomPos(atoms, k), p4 := AtomPos(atoms, l)
    if (!(IsObject(p1) && IsObject(p2) && IsObject(p3) && IsObject(p4)))
        return ""
    b1 := VecSub(p2, p1)
    b2 := VecSub(p3, p2)
    b3 := VecSub(p4, p3)
    lb2 := VecNorm(b2)
    if (lb2 = 0)
        return ""
    n1 := VecCross(b1, b2)
    n2 := VecCross(b2, b3)
    m1 := VecCross(n1, VecScale(b2, 1.0 / lb2))
    ln2 := VecNorm(n2)
    if (ln2 = 0)
        return ""
    x := VecDot(n1, n2)
    y := VecDot(m1, n2)
    return RadToDeg(ATan2(y, x))
}

VecSub(a, b)
{
    return [a[1] - b[1], a[2] - b[2], a[3] - b[3]]
}

VecDot(a, b)
{
    return a[1]*b[1] + a[2]*b[2] + a[3]*b[3]
}

VecCross(a, b)
{
    return [a[2]*b[3] - a[3]*b[2], a[3]*b[1] - a[1]*b[3], a[1]*b[2] - a[2]*b[1]]
}

VecScale(a, s)
{
    return [a[1]*s, a[2]*s, a[3]*s]
}

VecNorm(a)
{
    return Sqrt(VecDot(a, a))
}

RadToDeg(r)
{
    return r * 180.0 / ACos(-1.0)
}

; ATan2 is not a built-in in AHK v2, so provide it here
ATan2(y, x)
{
    static pi := ACos(-1.0)
    if (y = 0 && x = 0)
        return 0
    if (x > 0)
        return ATan(y / x)
    if (x < 0)
        return (y >= 0) ? ATan(y / x) + pi : ATan(y / x) - pi
    return (y > 0) ? pi / 2 : -pi / 2
}