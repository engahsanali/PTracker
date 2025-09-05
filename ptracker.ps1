# Project Folder Creator ++ — PS 5.1 (enhanced dashboard, tags, saved searches, notes sidebar, brace-safe)

$ErrorActionPreference='Stop'

# ===== USER SETTINGS =====
$BasePath = "C:\Users\N064176\OneDrive - Western Power\oneDrive-Ahsan-WP\Projects"
$EnableSerialPrefix = 'N'   # 'Y' to prefix with serial at the parent folder, anything else disables
$TemplatePath = "C:\Users\N064176\OneDrive - Western Power\Templates and Guides\Work Instructions\Customer coordination\Verification Checklist.xlsm"
$TemplateFileName = 'Verification Checklist.xlsm'
$KillProcessNames = @('EXCEL','WINWORD','POWERPNT')
# Keep-Awake / Idle detection
$KA_EnabledAtStart = $true
$KA_MinDelaySec = 15
$KA_MaxDelaySec = 30
$KA_MousePixels = 1
$KA_SendF24 = $true
$IdlePauseSec = 90     # if idle >= this, auto-pause time tracking; resume on activity
# Logging
$MaxLogMB = 5
$LoggingOnAtStart = $true
# =========================

Add-Type -AssemblyName System.Windows.Forms, System.Drawing | Out-Null
Add-Type -AssemblyName Microsoft.VisualBasic | Out-Null

# ---------- helpers ----------
function _TS([datetime]$dt=(Get-Date)){ $dt.ToString('yyyy-MM-dd HH:mm:ss') }
function _N([string]$s){ if($null -eq $s){''} else { $s.Trim() } }
function _LC([string]$s){ if($null -eq $s){''} else { $s.Trim().ToLower() } }
function _Esc([string]$s){ if($null -eq $s){''} else { $s -replace '\{','{{' -replace '\}','}}' } }
function _Pad2([int]$n){ ('{0:D2}' -f $n) }

# -------- single 5MB log (no duplicates) --------
$Global:LoggingEnabled = $LoggingOnAtStart
$Desktop = [Environment]::GetFolderPath('Desktop')
$Global:LogPath = Join-Path $Desktop 'ProjectFolderCreator.log'
function Ensure-Log{
  if(-not $Global:LoggingEnabled){ return }
  if(-not (Test-Path $Global:LogPath)){
    New-Item -ItemType File -Path $Global:LogPath -Force | Out-Null
    return
  }
  $len = ([IO.FileInfo]$Global:LogPath).Length/1MB
  if($len -ge $MaxLogMB){
    Set-Content -Path $Global:LogPath -Value ("[" + (_TS) + "] LOG TRUNCATED (> " + $MaxLogMB + " MB)") -Encoding UTF8
  }
}
function Write-Log([string]$m,[Exception]$ex=$null,[string]$tag='APP'){
  try{
    if(-not $Global:LoggingEnabled){return}; Ensure-Log
    $safe = _Esc $m
    $line = "[" + (_TS) + "] [" + $tag + "] " + $safe
    if($ex){ $line += " :: " + (_Esc $ex.Message) }
    Add-Content -Path $Global:LogPath -Value $line -Encoding UTF8
  }catch{}
}
function Open-Logs{ if(-not (Test-Path $Global:LogPath)){ Ensure-Log } ; Start-Process notepad.exe -ArgumentList $Global:LogPath }

# ---- base path autodetect; move log under BasePath as canonical ----
if(-not $BasePath -or -not (Test-Path $BasePath)){
  $u = Get-ChildItem C:\Users -Directory | ? Name -match '^N\d{6}$' | Select -First 1 -Expand Name
  if(-not $u){ throw "Could not auto-detect user." }
  $od = Get-ChildItem "C:\Users\$u" -Directory | ? Name -like '*OneDrive*' | Select -First 1
  if(-not $od){ throw "No OneDrive found" }
  $BasePath = Join-Path $od.FullName 'Projects'
  if(-not (Test-Path $BasePath)){ New-Item -ItemType Directory -Path $BasePath | Out-Null }
}
try{
  $newLog = Join-Path $BasePath 'ProjectFolderCreator.log'
  if((Test-Path $Global:LogPath) -and ($Global:LogPath -ne $newLog)){
    try{
      Get-Content $Global:LogPath -ErrorAction SilentlyContinue | Add-Content $newLog -Encoding UTF8
      Remove-Item $Global:LogPath -ErrorAction SilentlyContinue
    }catch{}
  }
  $Global:LogPath=$newLog; Ensure-Log; Write-Log ("Log target: " + $Global:LogPath)
}catch{}

# -------------- metadata & storage --------------
function _MetaRoot(){ $p=Join-Path $BasePath '.pfcmeta\projects'; if(-not (Test-Path $p)){ New-Item -ItemType Directory -Path $p -Force | Out-Null } ; $p }
function _DataRoot(){ $p=Join-Path $BasePath '.pfcmeta\data'; if(-not (Test-Path $p)){ New-Item -ItemType Directory -Path $p -Force | Out-Null } ; $p }
function _Safe([string]$s){ if($null -eq $s){'project'} else { $s -replace '[^\w\-]','_' } }
function _MetaPath([string]$proj){ Join-Path (_MetaRoot) ("{0}.json" -f (_Safe $proj)) }
$SavedSearchPath = Join-Path (_DataRoot) 'saved-searches.json'
if(-not (Test-Path $SavedSearchPath)){ '[]' | Set-Content -Path $SavedSearchPath -Encoding UTF8 }

$AllStatuses=@('New','Review','Sent for more info','Accepted','Auto Expired','Manually Rejected','Other')

function Load-Meta([string]$proj){
  $p=_MetaPath $proj
  if(Test-Path $p){ try{ return (Get-Content -Raw -LiteralPath $p | ConvertFrom-Json) }catch{} }
  [PSCustomObject]@{
    project=$proj
    createdAtUtc=(Get-Date).ToUniversalTime()
    currentPath=$null
    currentStatus='New'
    statusHistory=@(@{ts=(Get-Date).ToUniversalTime(); status='New'; note=''})
    notes=@()         # each note: {ts, text}
    sessions=@()      # {startUtc,endUtc,elapsedSec,workedSec,activeOnly}
    timeline=@()      # {ts,type,text} -- includes notes and statuses
    tags=@()          # array of strings
  }
}
function Save-Meta($m){
  try{ $m | ConvertTo-Json -Depth 7 | Set-Content -LiteralPath (_MetaPath $m.project) -Encoding UTF8 }catch{ Write-Log "Meta save failed" $_ 'META' }
}
function Total-WorkedSec($m){ $t=0; foreach($s in $m.sessions){ if($s.PSObject.Properties.Name -contains 'workedSec'){ $t+=[int]$s.workedSec } else { $t+=[int]$s.elapsedSec } } ; $t }

# -------------- folder helpers & indexer --------------
$ProjectFolderNamePattern = '^(?:\d+\.\s+)?(.+?)\s*-\s*Verification Documents$'

function Ensure-ChildFolders([string]$parent,[string]$proj){
  foreach($n in @("$proj - Customer Documents","$proj - WP Documents")){
    $d=Join-Path $parent $n; if(-not (Test-Path $d)){ New-Item -ItemType Directory -Path $d | Out-Null }
  }
}
function Next-Serial([string]$parent){
  $n=@(); Get-ChildItem -Path $parent -Directory | % { if($_.Name -match '^(\d+)\.\s+'){ $n+=[int]$matches[1] } }
  if($n.Count -eq 0){1}else{([int]($n|Measure-Object -Maximum).Maximum)+1}
}
function Resolve-TemplateFile(){
  if($TemplatePath -and (Test-Path $TemplatePath)){ return $TemplatePath }
  foreach($r in @($BasePath, [Environment]::GetFolderPath('Desktop'))){
    try{
      $f=Get-ChildItem -Path $r -File -EA SilentlyContinue |
         ? { $_.Name -ieq $TemplateFileName -or $_.Name -imatch 'verification.*checklist.*\.xlsm$' } |
         Select -First 1
      if($f){ return $f.FullName }
    }catch{}
  }
  $null
}
function Build-Meta-FromFolder([IO.DirectoryInfo]$dir){
  if(-not $dir -or -not $dir.Name){ return $null }
  if(-not ($dir.Name -match $ProjectFolderNamePattern)){ return $null }
  $proj=($matches[1]).Trim()
  $m=Load-Meta $proj
  $m.project=$proj; $m.currentPath=$dir.FullName
  if(-not $m.createdAtUtc){ $m.createdAtUtc=$dir.CreationTimeUtc }
  if(-not $m.statusHistory){ $m.statusHistory=@(@{ts=(Get-Date).ToUniversalTime(); status='New'; note=''}) }
  Save-Meta $m; return $m
}
function Rebuild-Index([switch]$Quiet){
  $dirs=Get-ChildItem -Path $BasePath -Recurse -Directory -EA SilentlyContinue
  $c=0; foreach($d in $dirs){ if($d.Name -match $ProjectFolderNamePattern){ $null=Build-Meta-FromFolder $d; $c++ } }
  if(-not $Quiet){ Write-Log ("Index refresh complete: " + $c + " project folders indexed") }
}

# -------------- lockers & move --------------
if(-not ("RestartManager.NativeMethods" -as [type])){
Add-Type -Language CSharp @"
using System; using System.Runtime.InteropServices; using System.Text; using FILETIME = System.Runtime.InteropServices.ComTypes.FILETIME;
namespace RestartManager {
  public enum RM_APP_TYPE { RmUnknownApp=0, RmMainWindow=1, RmOtherWindow=2, RmService=3, RmExplorer=4, RmConsole=5, RmCritical=1000 }
  [StructLayout(LayoutKind.Sequential)] public struct RM_UNIQUE_PROCESS { public int dwProcessId; public FILETIME ProcessStartTime; }
  [StructLayout(LayoutKind.Sequential, CharSet=CharSet.Unicode)] public struct RM_PROCESS_INFO {
    public RM_UNIQUE_PROCESS Process; [MarshalAs(UnmanagedType.ByValTStr, SizeConst=256)] public string strAppName; [MarshalAs(UnmanagedType.ByValTStr, SizeConst=64)] public string strServiceShortName; public RM_APP_TYPE ApplicationType; public uint AppStatus; public uint TSSessionId; [MarshalAs(UnmanagedType.Bool)] public bool bRestartable; }
  public static class NativeMethods {
    [DllImport("rstrtmgr.dll", CharSet=CharSet.Unicode)] public static extern int RmStartSession(out uint pSessionHandle, int dwSessionFlags, StringBuilder strSessionKey);
    [DllImport("rstrtmgr.dll", CharSet=CharSet.Unicode)] public static extern int RmRegisterResources(uint pSessionHandle, uint nFiles, string[] rgsFilenames, uint nApplications, IntPtr rgApplications, uint nServices, string[] rgsServiceNames);
    [DllImport("rstrtmgr.dll")] public static extern int RmGetList(uint dwSessionHandle, out uint pnProcInfoNeeded, ref uint pnProcInfo, [In, Out] RM_PROCESS_INFO[] rgAffectedApps, ref uint lpdwRebootReasons);
    [DllImport("rstrtmgr.dll")] public static extern int RmEndSession(uint pSessionHandle);
  }
}
"@ | Out-Null
}
function Close-Explorer([string]$p){
  try{
    $full=[IO.Path]::GetFullPath($p).TrimEnd('\'); $sh=New-Object -ComObject Shell.Application
    foreach($w in $sh.Windows()){ try{
      $fp=$w.Document.Folder.Self.Path
      if($fp -and ([IO.Path]::GetFullPath($fp).StartsWith($full,[StringComparison]::OrdinalIgnoreCase))){ $w.Quit() }
    }catch{} }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($sh)
  }catch{}
}
function Kill-Lockers([string]$folder){
  foreach($n in $KillProcessNames){ try{ Get-Process -Name $n -EA SilentlyContinue | Stop-Process -Force -EA SilentlyContinue }catch{} }
  try{
    $files=Get-ChildItem -LiteralPath $folder -Recurse -File -EA SilentlyContinue | % FullName | Select -Unique
    if(-not $files){ return }
    [uint32]$h=0; $k=[System.Text.StringBuilder]::new(256)
    [RestartManager.NativeMethods]::RmStartSession([ref]$h,0,$k)|Out-Null
    try{
      [RestartManager.NativeMethods]::RmRegisterResources($h,[uint32]$files.Count,$files,0,[IntPtr]::Zero,0,$null)|Out-Null
      [uint32]$need=0; [uint32]$have=0; [uint32]$reasons=0
      $rc=[RestartManager.NativeMethods]::RmGetList($h,[ref]$need,[ref]$have,$null,[ref]$reasons)
      if($rc -eq 234){
        $pi=New-Object RestartManager.RM_PROCESS_INFO[] $need; $have=$need
        $rc=[RestartManager.NativeMethods]::RmGetList($h,[ref]$need,[ref]$have,$pi,[ref]$reasons)
        if($rc -eq 0){
          $pids=@(); for($i=0;$i -lt $have;$i++){ $id=$pi[$i].Process.dwProcessId; if($id -and $id -ne $PID){ $pids+=$id } }
          $pids=$pids|Select -Unique; foreach($pid in $pids){ try{ Stop-Process -Id $pid -Force -EA SilentlyContinue }catch{} }
        }
      }
    } finally { [RestartManager.NativeMethods]::RmEndSession($h)|Out-Null }
  }catch{}
  Close-Explorer $folder
}
function Dir-Stats([string]$p){ if(-not (Test-Path $p)){return [pscustomobject]@{Files=0;Bytes=0}}; $f=Get-ChildItem -Recurse -File -LiteralPath $p -EA SilentlyContinue; [pscustomobject]@{Files=@($f).Count; Bytes=($f|Measure-Object -Sum Length).Sum} }
function Move-Tree([string]$src,[string]$destParent){
  $leaf=Split-Path $src -Leaf; $dst=Join-Path $destParent $leaf
  if(Test-Path $dst){ throw "Destination already has '$leaf'." }
  New-Item -ItemType Directory -Path $dst | Out-Null
  $args=@("""$src""","""$dst""","/E","/COPY:DAT","/DCOPY:DAT","/R:2","/W:1","/NFL","/NDL","/NP","/XJ")
  $rob=Start-Process robocopy.exe -ArgumentList $args -Wait -PassThru
  if($rob.ExitCode -ge 8){ throw "Robocopy failed ($($rob.ExitCode))" }
  $s=Dir-Stats $src; $d=Dir-Stats $dst
  if($d.Files -lt $s.Files -or $d.Bytes -lt $s.Bytes){ throw "Verification failed: destination smaller" }
  Remove-Item -LiteralPath $src -Recurse -Force -EA Stop
  $dst
}

# -------------- Idle detection (pause/resume) --------------
if(-not ("IdleNative" -as [type])){
Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class IdleNative {
  [StructLayout(LayoutKind.Sequential)] public struct LASTINPUTINFO { public uint cbSize; public uint dwTime; }
  [DllImport("user32.dll")] public static extern bool GetLastInputInfo(ref LASTINPUTINFO plii);
  public static uint GetIdleMillis() {
    var lii = new LASTINPUTINFO(); lii.cbSize=(uint)System.Runtime.InteropServices.Marshal.SizeOf(lii);
    GetLastInputInfo(ref lii);
    uint tick = (uint)Environment.TickCount;
    return tick - lii.dwTime;
  }
}
public static class K {
  [DllImport("user32.dll")] public static extern void keybd_event(byte vKey, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);
}
"@ | Out-Null
}
$Global:KA_Timer=$null
$Global:IdleTimer=$null
$Global:WasPaused=$false
function Start-KA{
  if($Global:KA_Timer){return}
  $Global:KA_Timer=New-Object Windows.Forms.Timer
  $Global:KA_Timer.Interval=(Get-Random -Minimum $KA_MinDelaySec -Maximum ($KA_MaxDelaySec+1))*1000
  $Global:KA_Timer.Add_Tick({
    try{
      $p=[Windows.Forms.Cursor]::Position
      [Windows.Forms.Cursor]::Position=New-Object Drawing.Point ($p.X+$KA_MousePixels),$p.Y
      Start-Sleep -Milliseconds 40
      [Windows.Forms.Cursor]::Position=$p
      if($KA_SendF24){ [K]::keybd_event(0x87,0,0,[UIntPtr]::Zero); Start-Sleep -Milliseconds 20; [K]::keybd_event(0x87,0,2,[UIntPtr]::Zero) }
      $Global:KA_Timer.Interval=(Get-Random -Minimum $KA_MinDelaySec -Maximum ($KA_MaxDelaySec+1))*1000
    }catch{}
  })
  $Global:KA_Timer.Start(); Write-Log "Keep-Awake started"
}
function Stop-KA{ if($Global:KA_Timer){ $Global:KA_Timer.Stop(); $Global:KA_Timer.Dispose(); $Global:KA_Timer=$null; Write-Log "Keep-Awake stopped" } }
function Start-IdleWatch{
  if($Global:IdleTimer){return}
  $Global:IdleTimer=New-Object Windows.Forms.Timer
  $Global:IdleTimer.Interval=1000
  $Global:IdleTimer.Add_Tick({
    try{
      $idleMs=[IdleNative]::GetIdleMillis()
      if(($idleMs -ge ($IdlePauseSec*1000)) -and ($Global:Current -ne $null) -and (-not $Global:WasPaused)){
        $Global:WasPaused=$true
        $Global:PauseStamp=(Get-Date).ToUniversalTime()
        Write-Log "Auto-pause (idle)"
      }
      if(($idleMs -lt 1500) -and $Global:WasPaused){
        $Global:WasPaused=$false
        Write-Log "Auto-resume (activity)"
      }
    }catch{}
  })
  $Global:IdleTimer.Start()
}
function Stop-IdleWatch{ if($Global:IdleTimer){ $Global:IdleTimer.Stop(); $Global:IdleTimer.Dispose(); $Global:IdleTimer=$null } }

# -------------- GUI --------------
[Windows.Forms.Application]::EnableVisualStyles()
$form=New-Object Windows.Forms.Form
$form.Text="Project Folder Creator ++"
$form.StartPosition='CenterScreen'
$form.AutoScaleMode='Font'
$form.MinimumSize=New-Object Drawing.Size(1120,680)
$form.Size=New-Object Drawing.Size(1220,720)

# Top user info
$envUser = $env:USERNAME
$nn = (Get-ChildItem C:\Users -Directory | ? FullName -like "*\$envUser" | Select -First 1).Name
$lblUser=New-Object Windows.Forms.Label
$lblUser.Text="User: $envUser ($nn)"
$lblUser.AutoSize=$true; $lblUser.Location='10,5'

# Row 1
$lblProj=New-Object Windows.Forms.Label; $lblProj.Text='Project #:'; $lblProj.Location='10,30'; $lblProj.AutoSize=$true; $lblProj.Anchor='Top,Left'
$txtProj=New-Object Windows.Forms.TextBox; $txtProj.Location='90,28'; $txtProj.Width=200; $txtProj.Anchor='Top,Left'
$btnMgr=New-Object Windows.Forms.Button; $btnMgr.Text='Search / Manage'; $btnMgr.Location='300,26'; $btnMgr.Width=130; $btnMgr.Anchor='Top,Left'
$btnNew=New-Object Windows.Forms.Button; $btnNew.Text='New'; $btnNew.Location='440,26'; $btnNew.Width=70; $btnNew.Anchor='Top,Left'
$btnKill=New-Object Windows.Forms.Button; $btnKill.Text='Kill Lockers'; $btnKill.Location='520,26'; $btnKill.Width=95; $btnKill.Anchor='Top,Left'
$btnOpenBase=New-Object Windows.Forms.Button; $btnOpenBase.Text='Open BasePath'; $btnOpenBase.Location='620,26'; $btnOpenBase.Width=120; $btnOpenBase.Anchor='Top,Left'

# quick apps
$btnCalc=New-Object Windows.Forms.Button; $btnCalc.Text='Calc'; $btnCalc.Location='750,26'; $btnCalc.Width=55
$btnExcel=New-Object Windows.Forms.Button; $btnExcel.Text='Excel'; $btnExcel.Location='810,26'; $btnExcel.Width=55
$btnWord=New-Object Windows.Forms.Button; $btnWord.Text='Word'; $btnWord.Location='870,26'; $btnWord.Width=55
$btnOutlook=New-Object Windows.Forms.Button; $btnOutlook.Text='Outlook'; $btnOutlook.Location='930,26'; $btnOutlook.Width=65

# Row 2
$lblDest=New-Object Windows.Forms.Label; $lblDest.Text='Move to:'; $lblDest.Location='10,65'; $lblDest.AutoSize=$true; $lblDest.Anchor='Top,Left'
$cboDest=New-Object Windows.Forms.ComboBox; $cboDest.Location='90,60'; $cboDest.Width=460; $cboDest.DropDownStyle='DropDownList'; $cboDest.Anchor='Top,Left,Right'
$chkDry=New-Object Windows.Forms.CheckBox; $chkDry.Text='Dry-run'; $chkDry.Location='560,62'; $chkDry.AutoSize=$true; $chkDry.Anchor='Top,Left'
$btnMove=New-Object Windows.Forms.Button; $btnMove.Text='Move'; $btnMove.Location='630,60'; $btnMove.Width=70; $btnMove.Anchor='Top,Left'
$btnUndo=New-Object Windows.Forms.Button; $btnUndo.Text='Undo Last'; $btnUndo.Location='710,60'; $btnUndo.Width=90; $btnUndo.Anchor='Top,Left'
$btnSettings=New-Object Windows.Forms.Button; $btnSettings.Text='Settings'; $btnSettings.Location='810,60'; $btnSettings.Width=80

# Keep-Awake group
$grpKA=New-Object Windows.Forms.GroupBox; $grpKA.Text='Keep-Awake'; $grpKA.Location='10,95'; $grpKA.Size='450,110'; $grpKA.Anchor='Top,Left'
$chkKA=New-Object Windows.Forms.CheckBox; $chkKA.Text='Enabled'; $chkKA.Location='15,25'; $chkKA.AutoSize=$true
$lblKA=New-Object Windows.Forms.Label; $lblKA.Text='(The app will nudge while showing Idle until you act)'; $lblKA.Location='15,55'; $lblKA.AutoSize=$true
$grpKA.Controls.AddRange(@($chkKA,$lblKA))

# Options group
$grpOpt=New-Object Windows.Forms.GroupBox; $grpOpt.Text='Options'; $grpOpt.Location='470,95'; $grpOpt.Size='430,110'; $grpOpt.Anchor='Top,Left'
$chkLog=New-Object Windows.Forms.CheckBox; $chkLog.Text='Enable logging'; $chkLog.Location='15,30'; $chkLog.AutoSize=$true
$chkTpl=New-Object Windows.Forms.CheckBox; $chkTpl.Text='Copy template on New'; $chkTpl.Location='15,60'; $chkTpl.AutoSize=$true
$grpOpt.Controls.AddRange(@($chkLog,$chkTpl))

# Row buttons
$btnLogs=New-Object Windows.Forms.Button; $btnLogs.Text='Open Logs'; $btnLogs.Location='10,210'; $btnLogs.Anchor='Top,Left'
$btnSummary=New-Object Windows.Forms.Button; $btnSummary.Text='Summary (Today)'; $btnSummary.Location='110,210'; $btnSummary.Anchor='Top,Left'
$btnExport=New-Object Windows.Forms.Button; $btnExport.Text='Export Timesheet (CSV)'; $btnExport.Location='260,210'; $btnExport.Anchor='Top,Left'
$btnExit=New-Object Windows.Forms.Button; $btnExit.Text='Exit'; $btnExit.Location='440,210'; $btnExit.Anchor='Top,Left'

# Dashboard - bottom left (recent projects list)
$lblDash=New-Object Windows.Forms.Label; $lblDash.Text='Recent Projects'; $lblDash.Location='10,240'; $lblDash.AutoSize=$true
$lvDash=New-Object Windows.Forms.ListView
$lvDash.Location='10,260'; $lvDash.Size='740,380'; $lvDash.View='Details'; $lvDash.FullRowSelect=$true; $lvDash.Anchor='Top,Left,Bottom'
[void]$lvDash.Columns.Add('Project',140)
[void]$lvDash.Columns.Add('Status',160)
[void]$lvDash.Columns.Add('Last Opened',140)
[void]$lvDash.Columns.Add('Total (hh:mm)',120)
[void]$lvDash.Columns.Add('Tags',160)

# Notes sidebar (right)
$grpNotes=New-Object Windows.Forms.GroupBox; $grpNotes.Text='Project Notes'; $grpNotes.Location='760,240'; $grpNotes.Size='430,400'; $grpNotes.Anchor='Top,Right,Bottom'
$lstNotes=New-Object Windows.Forms.ListBox; $lstNotes.Location='10,20'; $lstNotes.Size='410,300'; $lstNotes.Anchor='Top,Left,Right,Bottom'
$btnNAdd=New-Object Windows.Forms.Button; $btnNAdd.Text='Add'; $btnNAdd.Location='10,330'
$btnNEdit=New-Object Windows.Forms.Button; $btnNEdit.Text='Edit'; $btnNEdit.Location='70,330'
$btnNDel=New-Object Windows.Forms.Button; $btnNDel.Text='Delete'; $btnNDel.Location='130,330'
$grpNotes.Controls.AddRange(@($lstNotes,$btnNAdd,$btnNEdit,$btnNDel))

# Log area (top-right small)
$lstLog=New-Object Windows.Forms.ListBox; $lstLog.Location='910,95'; $lstLog.Size='280,110'; $lstLog.Anchor='Top,Right'
function UI([string]$m){ $lstLog.Items.Insert(0,("[" + (Get-Date).ToString('HH:mm:ss') + "] " + (_Esc $m))) ; Write-Log $m }

$form.Controls.AddRange(@(
  $lblUser,
  $lblProj,$txtProj,$btnMgr,$btnNew,$btnKill,$btnOpenBase,$btnCalc,$btnExcel,$btnWord,$btnOutlook,
  $lblDest,$cboDest,$chkDry,$btnMove,$btnUndo,$btnSettings,
  $grpKA,$grpOpt,$btnLogs,$btnSummary,$btnExport,$btnExit,
  $lblDash,$lvDash,$grpNotes,$lstLog
))

$Global:Current=$null; $Global:LastMove=$null; $Global:SessionStart=$null; $Global:SessionWorked=0; $Global:ActiveOnly=$true
function Enable-Move([bool]$on){ $cboDest.Enabled=$on; $chkDry.Enabled=$on; $btnMove.Enabled=$on; $btnUndo.Enabled=$on; $btnExport.Enabled=$on; $lstNotes.Enabled=$on; $btnNAdd.Enabled=$on; $btnNEdit.Enabled=$on; $btnNDel.Enabled=$on }
function Load-Destinations{
  $cboDest.Items.Clear()
  $rootLeaf=Split-Path $BasePath -Leaf
  [void]$cboDest.Items.Add("[ROOT] " + $rootLeaf + " — " + $BasePath + "|" + $BasePath)
  Get-ChildItem -Path $BasePath -Directory -EA SilentlyContinue | Sort-Object Name | %{
    [void]$cboDest.Items.Add( $_.Name + " — " + $_.FullName + "|" + $_.FullName )
  }
  if($cboDest.Items.Count -gt 0){ $cboDest.SelectedIndex=0 }
}
Enable-Move $false; Load-Destinations

# ----- dashboard loader (recent) -----
function Load-Dashboard{
  $lvDash.Items.Clear()
  $files=Get-ChildItem -File -Path (_MetaRoot) -Filter *.json -EA SilentlyContinue
  $items=@()
  foreach($f in $files){
    try{ $m=Get-Content -Raw -LiteralPath $f.FullName | ConvertFrom-Json }catch{ continue }
    $last=($m.timeline | Sort-Object ts -Descending | Select -First 1).ts
    if(-not $last){ $last=$m.createdAtUtc }
    $sec=Total-WorkedSec $m; $hh=[int]($sec/3600); $mm=[int](($sec%3600)/60)
    $it=New-Object Windows.Forms.ListViewItem $m.project
    [void]$it.SubItems.Add($m.currentStatus)
    [void]$it.SubItems.Add(([datetime]$last).ToLocalTime().ToString('yyyy-MM-dd HH:mm'))
    [void]$it.SubItems.Add((_Pad2 $hh)+":" + (_Pad2 $mm))
    [void]$it.SubItems.Add(($m.tags -join ', '))
    [void]$lvDash.Items.Add($it)
  }
  # most recent first
  $sorted=@($lvDash.Items) | Sort-Object { Get-Date $_.SubItems[2].Text } -Descending
  $lvDash.Items.Clear(); foreach($i in $sorted | Select -First 18){ [void]$lvDash.Items.Add($i) }
}
Load-Dashboard

# notes sidebar helpers
function Refresh-Notes(){
  $lstNotes.Items.Clear()
  if(-not $Global:Current){ return }
  $m=Load-Meta $Global:Current.Proj
  foreach($n in ($m.notes|Sort-Object ts -Descending)){
    $lstNotes.Items.Add(([datetime]$n.ts).ToLocalTime().ToString('yyyy-MM-dd HH:mm') + " : " + (_Esc $n.text)) | Out-Null
  }
}

# Session timer (paused on idle)
$secTimer=New-Object Windows.Forms.Timer; $secTimer.Interval=1000
$secTimer.Add_Tick({ if($Global:Current -and $Global:SessionStart -and (-not $Global:WasPaused)){ $Global:SessionWorked++ } })
$secTimer.Start()

# -------- small dialogs --------
function Dialog-Input([string]$title,[string]$prompt,[string]$prefill=''){
  $d=New-Object Windows.Forms.Form; $d.Text=$title; $d.AutoScaleMode='Font'; $d.Size='520,190'; $d.StartPosition='CenterParent'
  $l=New-Object Windows.Forms.Label; $l.Text=$prompt; $l.Location='10,10'; $l.AutoSize=$true
  $t=New-Object Windows.Forms.TextBox; $t.Location='10,35'; $t.Width=480; $t.Text=$prefill; $t.Anchor='Top,Left,Right'
  $ok=New-Object Windows.Forms.Button; $ok.Text='OK'; $ok.Location='10,80'
  $ca=New-Object Windows.Forms.Button; $ca.Text='Cancel'; $ca.Location='90,80'
  $d.Controls.AddRange(@($l,$t,$ok,$ca))
  $out=$null; $ok.Add_Click({ $out=$t.Text; $d.Close() }); $ca.Add_Click({ $d.Close() })
  $t.Add_KeyDown({ if($_.KeyCode -eq 'Enter'){ $ok.PerformClick(); $_.SuppressKeyPress=$true } })
  [void]$d.ShowDialog($form); $out
}

# -------- settings dialog --------
function Show-Settings{
  $s=New-Object Windows.Forms.Form; $s.Text='Settings'; $s.AutoScaleMode='Font'; $s.Size='420,260'; $s.StartPosition='CenterParent'
  $chk1=New-Object Windows.Forms.CheckBox; $chk1.Text='Enable logging'; $chk1.Checked=$Global:LoggingEnabled; $chk1.Location='20,20'; $chk1.AutoSize=$true
  $chk2=New-Object Windows.Forms.CheckBox; $chk2.Text='Copy template on New'; $chk2.Checked=$chkTpl.Checked; $chk2.Location='20,50'; $chk2.AutoSize=$true
  $lblMB=New-Object Windows.Forms.Label; $lblMB.Text='Max Log Size (MB):'; $lblMB.Location='20,80'; $lblMB.AutoSize=$true
  $numMB=New-Object Windows.Forms.NumericUpDown; $numMB.Location='150,78'; $numMB.Minimum=1; $numMB.Maximum=50; $numMB.Value=$MaxLogMB
  $ok=New-Object Windows.Forms.Button; $ok.Text='OK'; $ok.Location='20,120'
  $ok.Add_Click({
    $Global:LoggingEnabled=$chk1.Checked; $chkTpl.Checked=$chk2.Checked; $script:MaxLogMB=[int]$numMB.Value
    $s.Close()
  })
  $s.Controls.AddRange(@($chk1,$chk2,$lblMB,$numMB,$ok))
  [void]$s.ShowDialog($form)
}

# -------- Search/Manage with tags & saved searches --------
function Show-Manager([string]$initial){
  Rebuild-Index -Quiet
  $dlg=New-Object Windows.Forms.Form; $dlg.Text='Search / Manage'; $dlg.AutoScaleMode='Font'; $dlg.Size='1180,680'; $dlg.StartPosition='CenterParent'
  $lbl=New-Object Windows.Forms.Label; $lbl.Text='Find:'; $lbl.Location='10,12'; $lbl.AutoSize=$true
  $tb=New-Object Windows.Forms.TextBox; $tb.Location='50,10'; $tb.Width=230; $tb.Text=_N $initial
  $ls=New-Object Windows.Forms.Label; $ls.Text='Status:'; $ls.Location='290,12'; $ls.AutoSize=$true
  $cb=New-Object Windows.Forms.ComboBox; $cb.Location='340,10'; $cb.Width=150; $cb.DropDownStyle='DropDownList'
  [void]$cb.Items.Add('All'); foreach($s in $AllStatuses){ [void]$cb.Items.Add($s) } ; $cb.SelectedIndex=0
  $lt=New-Object Windows.Forms.Label; $lt.Text='Tag contains:'; $lt.Location='500,12'; $lt.AutoSize=$true
  $tbTag=New-Object Windows.Forms.TextBox; $tbTag.Location='580,10'; $tbTag.Width=140

  $lblSS=New-Object Windows.Forms.Label; $lblSS.Text='Saved:'; $lblSS.Location='730,12'; $lblSS.AutoSize=$true
  $cbSS=New-Object Windows.Forms.ComboBox; $cbSS.Location='780,10'; $cbSS.Width=210; $cbSS.DropDownStyle='DropDownList'
  function Load-SS{
    $json=Get-Content -Raw -Path $SavedSearchPath -EA SilentlyContinue
    if(-not $json){ $json='[]' }
    $script:Saved=($json | ConvertFrom-Json)
    $cbSS.Items.Clear(); [void]$cbSS.Items.Add('<none>')
    foreach($x in $script:Saved){ [void]$cbSS.Items.Add($x.name) }
    $cbSS.SelectedIndex=0
  }
  Load-SS
  $btnSave=New-Object Windows.Forms.Button; $btnSave.Text='Save'; $btnSave.Location='1000,8'
  $btnR=New-Object Windows.Forms.Button; $btnR.Text='Refresh'; $btnR.Location='1060,8'

  $lv=New-Object Windows.Forms.ListView; $lv.Location='10,40'; $lv.Size='1140,560'; $lv.View='Details'; $lv.FullRowSelect=$true; $lv.Anchor='Top,Left,Right,Bottom'
  [void]$lv.Columns.Add('Project',140); [void]$lv.Columns.Add('Status',150); [void]$lv.Columns.Add('Total (hh:mm)',120); [void]$lv.Columns.Add('Created',120); [void]$lv.Columns.Add('Tags',220); [void]$lv.Columns.Add('Folder',360)

  $bOpen=New-Object Windows.Forms.Button; $bOpen.Text='Open'; $bOpen.Location='10,606'
  $bMove=New-Object Windows.Forms.Button; $bMove.Text='Move…'; $bMove.Location='90,606'
  $bDel=New-Object Windows.Forms.Button; $bDel.Text='Delete (Recycle)'; $bDel.Location='170,606'; $bDel.Width=130
  $bKill=New-Object Windows.Forms.Button; $bKill.Text='Kill Lockers'; $bKill.Location='310,606'
  $bStat=New-Object Windows.Forms.Button; $bStat.Text='Set Status'; $bStat.Location='410,606'
  $bNote=New-Object Windows.Forms.Button; $bNote.Text='Add Note'; $bNote.Location='510,606'
  $bTime=New-Object Windows.Forms.Button; $bTime.Text='Timeline'; $bTime.Location='600,606'
  $bTags=New-Object Windows.Forms.Button; $bTags.Text='Edit Tags'; $bTags.Location='690,606'
  $bClose=New-Object Windows.Forms.Button; $bClose.Text='Close'; $bClose.Location='1060,606'

  $dlg.Controls.AddRange(@($lbl,$tb,$ls,$cb,$lt,$tbTag,$lblSS,$cbSS,$btnSave,$btnR,$lv,$bOpen,$bMove,$bDel,$bKill,$bStat,$bNote,$bTime,$bTags,$bClose))
  $cbSS.Add_SelectedIndexChanged({
    if($cbSS.SelectedIndex -le 0){ return }
    $s=$script:Saved | ? name -eq $cbSS.SelectedItem
    if($s){
      $tb.Text=$s.find; $cb.Text=$s.status; $tbTag.Text=$s.tag
      $btnR.PerformClick()
    }
  })
  $btnSave.Add_Click({
    $name=Dialog-Input "Save search" "Name for saved search:"
    if([string]::IsNullOrWhiteSpace($name)){ return }
    $obj=[pscustomobject]@{name=$name; find=$tb.Text; status=$cb.Text; tag=$tbTag.Text}
    $arr=@()
    if(Test-Path $SavedSearchPath){ $arr = (Get-Content -Raw -Path $SavedSearchPath | ConvertFrom-Json) }
    $arr += $obj
    $arr | ConvertTo-Json -Depth 4 | Set-Content -Path $SavedSearchPath -Encoding UTF8
    Load-SS()
  })

  function Load-List{
    $lv.Items.Clear()
    $needle=_LC $tb.Text; $want=$cb.SelectedItem; $tagNeed=_LC $tbTag.Text
    $files=Get-ChildItem -File -Path (_MetaRoot) -Filter *.json -EA SilentlyContinue
    foreach($f in $files){
      try{ $m=Get-Content -Raw -LiteralPath $f.FullName | ConvertFrom-Json }catch{ continue }
      $proj=[string]$m.project; $path=[string]$m.currentPath
      if($needle){ $hay=_LC ($proj+' '+$path); if($hay.IndexOf($needle) -lt 0){ continue } }
      if(($want) -and ($want -ne 'All')){
        $ok=$false
        foreach($sh in $m.statusHistory){ if($sh.status -eq $want){ $ok=$true; break } }
        if(-not $ok){ continue }
      }
      if($tagNeed){
        $hit=$false; foreach($t in $m.tags){ if(_LC $t -like "*$tagNeed*"){ $hit=$true; break } }
        if(-not $hit){ continue }
      }
      $tot=Total-WorkedSec $m; $hh=[int]($tot/3600); $mm=[int](($tot%3600)/60)
      $it=New-Object Windows.Forms.ListViewItem $proj
      [void]$it.SubItems.Add($m.currentStatus)
      [void]$it.SubItems.Add((_Pad2 $hh)+":" + (_Pad2 $mm))
      [void]$it.SubItems.Add(([datetime]$m.createdAtUtc).ToLocalTime().ToString('yyyy-MM-dd'))
      [void]$it.SubItems.Add(($m.tags -join ', '))
      [void]$it.SubItems.Add($path)
      [void]$lv.Items.Add($it)
    }
  }

  $btnR.Add_Click({ Rebuild-Index -Quiet; Load-List })
  $tb.Add_KeyDown({ if($_.KeyCode -eq 'Enter'){ Load-List; $_.SuppressKeyPress=$true } })
  $lv.Add_DoubleClick({ if($lv.SelectedItems.Count -gt 0){ $bOpen.PerformClick() } })

  $bOpen.Add_Click({
    if($lv.SelectedItems.Count -eq 0){ return }
    $p=$lv.SelectedItems[0].Text; $m=Load-Meta $p
    if(-not $m.currentPath -or -not (Test-Path $m.currentPath)){ [Windows.Forms.MessageBox]::Show("Folder not found.","Open"); return }
    if($Global:Current){
      $mp=Load-Meta $Global:Current.Proj
      $st=$Global:SessionStart; $en=(Get-Date).ToUniversalTime()
      if($st){ $mp.sessions+=@{startUtc=$st;endUtc=$en;elapsedSec=[int]($en-$st).TotalSeconds;workedSec=$Global:SessionWorked;activeOnly=$Global:ActiveOnly}; $mp.timeline+=@{ts=$en;type='close';text='Closed via manager'}; Save-Meta $mp }
    }
    $Global:Current=@{Proj=$p; Folder=$m.currentPath}; $Global:SessionStart=(Get-Date).ToUniversalTime(); $Global:SessionWorked=0
    $m.timeline+=@{ts=(Get-Date).ToUniversalTime(); type='open'; text='Opened via manager'}; Save-Meta $m
    Enable-Move $true; $txtProj.Text=$p; Start-Process explorer.exe -ArgumentList $m.currentPath; $dlg.Close()
    Refresh-Notes; Load-Dashboard
  })
  $bMove.Add_Click({
    if($lv.SelectedItems.Count -eq 0){ return }; $p=$lv.SelectedItems[0].Text; $m=Load-Meta $p
    if(-not (Test-Path $m.currentPath)){ [Windows.Forms.MessageBox]::Show("Folder not found.","Move"); return }
    $fb=New-Object System.Windows.Forms.FolderBrowserDialog; $fb.Description="Pick destination parent"; $fb.SelectedPath=$BasePath
    if($fb.ShowDialog() -ne 'OK'){ return }
    try{
      $new=Move-Tree $m.currentPath $fb.SelectedPath
      $m.timeline+=@{ts=(Get-Date).ToUniversalTime();type='move';text=("Moved to " + $new)}; $m.currentPath=$new; Save-Meta $m; Load-List; UI ("Relocated: " + $p)
    }catch{ [Windows.Forms.MessageBox]::Show("Move failed: $($_.Exception.Message)","Move") }
  })
  $bDel.Add_Click({
    if($lv.SelectedItems.Count -eq 0){ return }; $p=$lv.SelectedItems[0].Text; $m=Load-Meta $p
    if(-not (Test-Path $m.currentPath)){ [Windows.Forms.MessageBox]::Show("Folder not found.","Delete"); return }
    $ok=[Windows.Forms.MessageBox]::Show("Send to Recycle Bin?`n`n"+$m.currentPath,"Delete",'YesNo','Warning'); if($ok -ne 'Yes'){ return }
    try{
      [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteDirectory($m.currentPath,'OnlyErrorDialogs','SendToRecycleBin')
      $m.timeline+=@{ts=(Get-Date).ToUniversalTime();type='delete';text='Sent to Recycle Bin'}; Save-Meta $m; Load-List; UI ("Deleted: " + $p)
    }catch{ [Windows.Forms.MessageBox]::Show("Delete failed: $($_.Exception.Message)","Delete") }
  })
  $bKill.Add_Click({ if($lv.SelectedItems.Count -gt 0){ $p=$lv.SelectedItems[0].Text; $m=Load-Meta $p; if($m.currentPath){ Kill-Lockers $m.currentPath; UI ("Lockers closed for " + $p) } } })
  $bStat.Add_Click({
    if($lv.SelectedItems.Count -eq 0){ return }; $p=$lv.SelectedItems[0].Text; $m=Load-Meta $p
    $ch=New-Object Windows.Forms.Form; $ch.Text=("Set Status — "+$p); $ch.AutoScaleMode='Font'; $ch.Size='360,180'; $ch.StartPosition='CenterParent'
    $cb2=New-Object Windows.Forms.ComboBox; $cb2.Location='10,10'; $cb2.Width=320; $cb2.DropDownStyle='DropDownList'; foreach($s in $AllStatuses){ [void]$cb2.Items.Add($s) }; $cb2.Text=$m.currentStatus
    $ok=New-Object Windows.Forms.Button; $ok.Text='OK'; $ok.Location='10,50'; $ca=New-Object Windows.Forms.Button; $ca.Text='Cancel'; $ca.Location='90,50'
    $ch.Controls.AddRange(@($cb2,$ok,$ca)); $chosen=$null; $ok.Add_Click({ $chosen=$cb2.Text; $ch.Close() }); $ca.Add_Click({ $ch.Close() }); [void]$ch.ShowDialog($dlg)
    if(-not $chosen){return}; $note=Dialog-Input "Optional note" "Note for '$chosen':"
    $m.currentStatus=$chosen; $m.statusHistory+=@{ts=(Get-Date).ToUniversalTime(); status=$chosen; note=($note -as [string])}; $m.timeline+=@{ts=(Get-Date).ToUniversalTime(); type='status'; text=("-> "+$chosen)}; Save-Meta $m; Load-List; Load-Dashboard
  })
  $bNote.Add_Click({
    if($lv.SelectedItems.Count -eq 0){ return }; $p=$lv.SelectedItems[0].Text; $m=Load-Meta $p
    $txt=Dialog-Input ("Add Note — "+$p) "Enter note:"; if([string]::IsNullOrWhiteSpace($txt)){ return }
    $m.notes+=@{ts=(Get-Date).ToUniversalTime(); text=$txt}; $m.timeline+=@{ts=(Get-Date).ToUniversalTime(); type='note'; text=$txt}; Save-Meta $m; UI ("Note added to "+$p); Load-Dashboard
  })
  $bTime.Add_Click({
    if($lv.SelectedItems.Count -eq 0){ return }
    $p=$lv.SelectedItems[0].Text; $m=Load-Meta $p
    # timeline viewer with delete
    $d=New-Object Windows.Forms.Form; $d.Text=("Timeline — "+$p); $d.AutoScaleMode='Font'; $d.Size='900,520'; $d.StartPosition='CenterParent'
    $lv2=New-Object Windows.Forms.ListView; $lv2.View='Details'; $lv2.FullRowSelect=$true; $lv2.Location='10,10'; $lv2.Size='860,420'; $lv2.Anchor='Top,Left,Right,Bottom'
    [void]$lv2.Columns.Add('When',180); [void]$lv2.Columns.Add('Type',160); [void]$lv2.Columns.Add('Text',500)
    function loadT{
      $lv2.Items.Clear()
      $events=@(); $events+=$m.timeline; foreach($s in $m.statusHistory){ $events+=[pscustomobject]@{ts=$s.ts; type=("status: "+$s.status); text=$s.note} }; foreach($n in $m.notes){ $events+=[pscustomobject]@{ts=$n.ts; type='note'; text=$n.text} }
      $events=$events|Sort-Object ts
      foreach($e in $events){
        $when=([datetime]$e.ts).ToLocalTime().ToString('yyyy-MM-dd HH:mm')
        $it=New-Object Windows.Forms.ListViewItem $when
        [void]$it.SubItems.Add($e.type); [void]$it.SubItems.Add((_Esc $e.text)); [void]$lv2.Items.Add($it)
      }
    }
    loadT
    $del=New-Object Windows.Forms.Button; $del.Text='Delete Selected'; $del.Location='10,440'
    $del.Add_Click({
      if($lv2.SelectedItems.Count -eq 0){ return }
      $when=$lv2.SelectedItems[0].Text
      $ts=[datetime]::ParseExact($when,'yyyy-MM-dd HH:mm',$null).ToUniversalTime()
      # remove matching items from timeline and notes with that ts
      $m.timeline = @($m.timeline | ? { ([datetime]$_.ts) -ne $ts })
      $m.notes    = @($m.notes    | ? { ([datetime]$_.ts) -ne $ts })
      Save-Meta $m; loadT
    })
    $cls=New-Object Windows.Forms.Button; $cls.Text='Close'; $cls.Location='140,440'; $cls.Add_Click({ $d.Close() })
    $d.Controls.AddRange(@($lv2,$del,$cls)); [void]$d.ShowDialog($dlg)
  })
  $bTags.Add_Click({
    if($lv.SelectedItems.Count -eq 0){ return }
    $p=$lv.SelectedItems[0].Text; $m=Load-Meta $p
    $txt=Dialog-Input ("Edit Tags — "+$p) "Comma separated:" ($m.tags -join ', ')
    if($txt -ne $null){ $m.tags=@(); foreach($t in $txt.Split(',').Trim()){ if($t){ $m.tags+=$t } } ; Save-Meta $m; Load-List; Load-Dashboard }
  })
  $bClose.Add_Click({ $dlg.Close() })
  $btnR.PerformClick()
  [void]$dlg.ShowDialog($form)
}

# ----- creation helper used by main -----
function Create-Project([string]$proj,[switch]$ShowExplorer){
  $proj=_N $proj; if(-not $proj){ return $null }
  Rebuild-Index -Quiet
  $meta=Get-ChildItem -File -Path (_MetaRoot) -Filter *.json -EA SilentlyContinue | % { try{ Get-Content -Raw -LiteralPath $_.FullName | ConvertFrom-Json }catch{} }
  if(($meta | ? { (_LC $_.project) -eq (_LC $proj) }).Count -gt 0){
    [Windows.Forms.MessageBox]::Show("Project already tracked: " + $proj,"New"); return $null
  }
  $fb=New-Object System.Windows.Forms.FolderBrowserDialog; $fb.Description="Select parent (Cancel => Projects\Assorted)"; $fb.SelectedPath=$BasePath
  $res=$fb.ShowDialog()
  $parent = if($res -eq 'OK'){ $fb.SelectedPath } else { $a=Join-Path $BasePath 'Assorted'; if(-not (Test-Path $a)){ New-Item -ItemType Directory -Path $a | Out-Null }; $a }
  $prefix = if($EnableSerialPrefix -eq 'Y'){ ('{0}. ' -f (Next-Serial $parent)) } else { "" }
  $name=$prefix + $proj + " - Verification Documents"; $target=Join-Path $parent $name
  if(Test-Path $target){ [Windows.Forms.MessageBox]::Show("Already exists: " + $target,"New"); return $null }
  New-Item -ItemType Directory -Path $target | Out-Null; Ensure-ChildFolders $target $proj
  if($chkTpl.Checked){ $src=Resolve-TemplateFile; if($src){ $wp=Join-Path $target ($proj + " - WP Documents"); Copy-Item -Path $src -Destination (Join-Path $wp ($proj + " - Verification Checklist.xlsm")) -Force } }
  $m=Load-Meta $proj; $m.createdAtUtc=(Get-Date).ToUniversalTime(); $m.currentPath=$target; $m.timeline+=@{ts=(Get-Date).ToUniversalTime(); type='create'; text='New project created'}; Save-Meta $m
  $Global:Current=@{Proj=$proj; Folder=$target}; $Global:SessionStart=(Get-Date).ToUniversalTime(); $Global:SessionWorked=0; Enable-Move $true; $txtProj.Text=$proj
  if($ShowExplorer){ Start-Process explorer.exe -ArgumentList $target }
  UI ("Created: " + $target)
  Refresh-Notes; Load-Dashboard
  return $target
}

# ---- MAIN handlers ----
$chkLog.Checked=$LoggingOnAtStart; $chkLog.Add_CheckedChanged({ $Global:LoggingEnabled=$chkLog.Checked; UI ("Logging: " + (if($Global:LoggingEnabled){"ON"}else{"OFF"})) })
$chkTpl.Checked=$true
$chkKA.Checked=$KA_EnabledAtStart; if($chkKA.Checked){ Start-KA }; $chkKA.Add_CheckedChanged({ if($chkKA.Checked){ Start-KA } else { Stop-KA } })
Start-IdleWatch

$btnOpenBase.Add_Click({ Start-Process explorer.exe -ArgumentList $BasePath })
$btnMgr.Add_Click({ Show-Manager (_N $txtProj.Text) })
$txtProj.Add_KeyDown({ if($_.KeyCode -eq 'Enter'){ Show-Manager (_N $txtProj.Text); $_.SuppressKeyPress=$true } })

$btnNew.Add_Click({ try{ Create-Project (_N $txtProj.Text) -ShowExplorer | Out-Null }catch{ [Windows.Forms.MessageBox]::Show("Create failed: $($_.Exception.Message)","New") } })
$btnKill.Add_Click({ try{ if(-not $Global:Current){ [Windows.Forms.MessageBox]::Show("Open or create a project first.","Kill"); return }; Kill-Lockers $Global:Current.Folder; UI "Lockers closed." }catch{ [Windows.Forms.MessageBox]::Show("Kill failed: $($_.Exception.Message)","Kill") } })
$btnSettings.Add_Click({ Show-Settings })

$btnMove.Add_Click({
  try{
    if(-not $Global:Current){ [Windows.Forms.MessageBox]::Show("Open or create a project first.","Move"); return }
    if($cboDest.SelectedItem -eq $null){ [Windows.Forms.MessageBox]::Show("Pick a destination.","Move"); return }
    $parts=$cboDest.SelectedItem.ToString().Split('|',2); $destParent=$parts[1]
    $src=$Global:Current.Folder; $leaf=Split-Path $src -Leaf; $dst=Join-Path $destParent $leaf
    if([IO.Path]::GetFullPath($src).TrimEnd('\') -ieq [IO.Path]::GetFullPath($dst).TrimEnd('\')){ [Windows.Forms.MessageBox]::Show("Cannot move into itself.","Move"); return }
    if(Test-Path $dst){ [Windows.Forms.MessageBox]::Show("Already exists at destination.","Move"); return }
    if($chkDry.Checked){
      $s=Dir-Stats $src
      [Windows.Forms.MessageBox]::Show("DRY RUN`nFrom:`n" + $src + "`nTo:`n" + $dst + "`nFiles: " + $s.Files + "  Bytes: " + $s.Bytes,"Move"); return
    }
    $new=Move-Tree $src $destParent; $m=Load-Meta $Global:Current.Proj; $m.timeline+=@{ts=(Get-Date).ToUniversalTime(); type='move'; text=("Moved to " + $new)}; $m.currentPath=$new; Save-Meta $m
    $Global:LastMove=@{From=$src; To=$new}; $Global:Current=@{Proj=$Global:Current.Proj; Folder=$new}; Start-Process explorer.exe -ArgumentList $new; UI ("Relocated: " + $src + " -> " + $new); Load-Dashboard
  }catch{ [Windows.Forms.MessageBox]::Show("Move failed: $($_.Exception.Message)","Move") }
})
$btnUndo.Add_Click({
  try{
    if(-not $Global:LastMove){ [Windows.Forms.MessageBox]::Show("No move to undo.","Undo"); return }
    $from=$Global:LastMove.To; $parent=(Split-Path $Global:LastMove.From -Parent); $new=Move-Tree $from $parent
    $m=Load-Meta $Global:Current.Proj; $m.currentPath=$new; $m.timeline+=@{ts=(Get-Date).ToUniversalTime(); type='move'; text='Undo move'}; Save-Meta $m
    $Global:Current=@{Proj=$Global:Current.Proj; Folder=$new}; $Global:LastMove=$null; Start-Process explorer.exe -ArgumentList $new; UI "Undo complete."; Load-Dashboard
  }catch{ [Windows.Forms.MessageBox]::Show("Undo failed: $($_.Exception.Message)","Undo") }
})

$btnLogs.Add_Click({ Open-Logs })
$btnSummary.Add_Click({
  try{
    $files=@(Get-ChildItem -File -Path (_MetaRoot) -Filter *.json -EA SilentlyContinue)
    $count=$files.Count
    [Windows.Forms.MessageBox]::Show("Projects tracked: " + $count + "`nUse Search/Manage to filter and open.","Summary")
  }catch{ [Windows.Forms.MessageBox]::Show("Summary failed: $($_.Exception.Message)","Summary") }
})
$btnExport.Add_Click({
  try{
    $files=Get-ChildItem -File -Path (_MetaRoot) -Filter *.json -EA SilentlyContinue
    if(-not $files){ [Windows.Forms.MessageBox]::Show("No projects tracked yet.","Export"); return }
    $out = Join-Path $BasePath ("Timesheet_" + (Get-Date).ToString('yyyyMMdd_HHmmss') + ".csv")
    "Project,Hours,Minutes" | Set-Content -Path $out -Encoding UTF8
    foreach($f in $files){
      try{ $m=Get-Content -Raw -LiteralPath $f.FullName | ConvertFrom-Json }catch{ continue }
      $sec=Total-WorkedSec $m; $hh=[int]($sec/3600); $mm=[int](($sec%3600)/60)
      $line = '"' + (($m.project) -replace '"','""') + '",' + $hh + ',' + $mm
      Add-Content -Path $out -Value $line -Encoding UTF8
    }
    Start-Process $out
  }catch{ [Windows.Forms.MessageBox]::Show("Export failed: $($_.Exception.Message)","Export") }
})
$btnExit.Add_Click({
  try{
    if($Global:Current){
      $m=Load-Meta $Global:Current.Proj
      $st=$Global:SessionStart; $en=(Get-Date).ToUniversalTime()
      if($st){ $m.sessions+=@{startUtc=$st;endUtc=$en;elapsedSec=[int]($en-$st).TotalSeconds;workedSec=$Global:SessionWorked;activeOnly=$Global:ActiveOnly}; $m.timeline+=@{ts=$en;type='close';text='Exit app'}; Save-Meta $m }
    }
  }catch{}
  Stop-KA; Stop-IdleWatch; $form.Close()
})

# notes sidebar actions
$btnNAdd.Add_Click({
  if(-not $Global:Current){ return }
  $p=$Global:Current.Proj; $t=Dialog-Input ("Add Note — "+$p) "Enter note:"; if([string]::IsNullOrWhiteSpace($t)){return}
  $m=Load-Meta $p; $m.notes+=@{ts=(Get-Date).ToUniversalTime(); text=$t}; $m.timeline+=@{ts=(Get-Date).ToUniversalTime(); type='note'; text=$t}; Save-Meta $m
  Refresh-Notes; Load-Dashboard
})
$btnNEdit.Add_Click({
  if(-not $Global:Current -or $lstNotes.SelectedIndex -lt 0){ return }
  $p=$Global:Current.Proj; $m=Load-Meta $p
  # parse selected line's timestamp
  $line=$lstNotes.SelectedItem; $tsLocal=[datetime]::ParseExact($line.Substring(0,16),'yyyy-MM-dd HH:mm',$null)
  $ts=$tsLocal.ToUniversalTime()
  $old=($m.notes | ? { ([datetime]$_.ts) -eq $ts } | Select -First 1)
  if($null -eq $old){ return }
  $new=Dialog-Input ("Edit Note — "+$p) "Update note:" $old.text
  if($new -ne $null){ $old.text=$new; Save-Meta $m; Refresh-Notes; Load-Dashboard }
})
$btnNDel.Add_Click({
  if(-not $Global:Current -or $lstNotes.SelectedIndex -lt 0){ return }
  $p=$Global:Current.Proj; $m=Load-Meta $p
  $line=$lstNotes.SelectedItem; $tsLocal=[datetime]::ParseExact($line.Substring(0,16),'yyyy-MM-dd HH:mm',$null)
  $ts=$tsLocal.ToUniversalTime()
  $m.notes = @($m.notes | ? { ([datetime]$_.ts) -ne $ts })
  $m.timeline = @($m.timeline | ? { ([datetime]$_.ts) -ne $ts -or $_.type -ne 'note' })
  Save-Meta $m; Refresh-Notes; Load-Dashboard
})

# quick apps
$btnCalc.Add_Click({ Start-Process calc.exe })
$btnExcel.Add_Click({ Start-Process excel.exe -ErrorAction SilentlyContinue })
$btnWord.Add_Click({ Start-Process winword.exe -ErrorAction SilentlyContinue })
$btnOutlook.Add_Click({
  if(-not (Get-Process -Name OUTLOOK -EA SilentlyContinue)){
    Start-Process outlook.exe -ArgumentList '/recycle' -ErrorAction SilentlyContinue
  }
})

# initialize
$chkKA.Checked=$KA_EnabledAtStart; if($chkKA.Checked){ Start-KA }
$form.Add_Shown({ $txtProj.Focus() })
[void]$form.ShowDialog()
