; This script was created using Pulover's Macro Creator
; www.macrocreator.com

#NoEnv
SetWorkingDir %A_ScriptDir%
CoordMode, Mouse, Window
SendMode Input
#SingleInstance Force
SetTitleMatchMode 2
#WinActivateForce
SetControlDelay 1
SetWinDelay 0
SetKeyDelay -1
SetMouseDelay -1
SetBatchLines -1


F3::
Macro1:
FileSelectFile, loginList, , C:\Users\rijul.kumar\Desktop\DGN Bots\Current Versions\DGN Planroom Log In.xlsx, Please select your Log In File
If !IsObject(Xl)
	Xl := ComObjCreate("Excel.Application")

Xl.Workbooks.Open(loginList)
endRow := Xl.columns("A").end(-4121).row
associations := %endRow% -1
texty := ""
textyRow := 2
While textyRow <= endRow
{
    name := 
    (LTrim
    Xl.range("A" . %textyRow%).value . " = " . (%textyRow%-1) . "
    
    "
    )
    texty .= name
    textyRow += 1
}
InputBox, looper, Where would you like to start your syncing process?, %texty%, , 500, 700, , , , , 14
startTime := A_TickCount
If !IsObject(ie)
	ie := ComObjCreate("InternetExplorer.Application")
ie.Visible := true
ie.Navigate("network2.construction.com")
IELoad(ie)
WinMaximize, construction.com - the construction industry marketplace - Internet Explorer
Sleep, 333
WinActivate, construction.com - the construction industry marketplace - Internet Explorer
Sleep, 333
DGN:
If looper <= %associations%
{
    Loop
    {
        rowNum := %looper%+1
        userName := Xl.range("A" . %rowNum%).value
        passWord := Xl.range("B" . %rowNum%).value
        SSID := Xl.range("C" . %rowNum%).value
        Send, %userName%  ; sign in
        CurrentKeyDelay := A_KeyDelay
        SetKeyDelay, 1
        SendEvent, {Tab}  ; sign in
        SetKeyDelay, %CurrentKeyDelay%
        Send, %passWord%  ; sign in
        CurrentKeyDelay := A_KeyDelay
        SetKeyDelay, 1
        SendEvent, {Enter}  ; sign in
        SetKeyDelay, %CurrentKeyDelay%
        Sleep, 500
        ie.Navigate("http://network2.construction.com/Home.aspx")
        IELoad(ie)
        ie.Navigate(SSID)
        IELoad(ie)
        first := 0
        projBool := ie.document.getElementByID("ctl00_contentPlaceHolderHeader_pcTop_listProjectCountText").InnerHTML
        If projBool !=
        {
            ie.document.getElementsByName("project-select-all")[0].Click("")
            ie.document.getElementsByName("ctl00$contentPlaceHolderHeader$pcTop$HeaderActions$btnprjresltAction")[0].Click("")
            ie.document.getElementByID("lnkViewProjects").Click("") := ""
        }
        Else
        {
            looper += 1
            ie.document.getElementByID("ctl00_ucHeader_lnk_SignOut").Click("") := ""
            IELoad(ie)
            Goto, DGN
        }
        Loop
        {
            CoordMode, Pixel, Window
            PixelSearch, FoundX, FoundY, 0, 0, %A_ScreenWidth%, %A_ScreenHeight%, 0x2DB98D, 0, Fast RGB
        }
        Until ErrorLevel = 0
        projNum := ie.document.getElementByID("ctl00_contentPlaceHolderHeader_rptPager_lblTotalPageCount").InnerText
        If projNum !=
        {
            loopNum := projNum
        }
        Else
        {
            loopNum := 1
        }
        Sleep, 1000
        While loopNum>=0
        {
            If loopNum > 0
            {
                2:
                ie.document.getElementsByClassName("syncRefreshText")[0].Click("")
                If ErrorLevel != 0
                {
                    Goto, 2
                }
                3:
                ie.document.getElementsByClassName("planRoomOkText")[1].Click("")
                If ErrorLevel != 0
                {
                    Goto, 3
                }
                4:
                ie.document.getElementByID("lnkTrackProjects").Click("") := ""
                If ErrorLevel != 0
                {
                    Goto, 4
                }
                5:
                ie.document.getElementsByClassName("trackCheck")[1].Click("")
                If ErrorLevel != 0
                {
                    Goto, 5
                }
                6:
                ie.document.getElementsByClassName("track-popup-submit")[0].Click("")
                If ErrorLevel != 0
                {
                    Goto, 6
                }
                7:
                ie.document.getElementsByClassName("syncRefreshText")[0].Focus("")
                If ErrorLevel != 0
                {
                    Goto, 7
                }
                Else
                {
                    If loopNum > 1
                    {
                        ie.document.getElementByID("ctl00_contentPlaceHolderHeader_rptPager_lblNext").Click("") := ""
                        IELoad(ie)
                    }
                }
                loopNum -= 1
            }
            Else
            {
                ie.document.getElementByID("ctl00_ucHeader_lnk_SignOut").Click("") := ""
                IELoad(ie)
                Sleep, 500
                looper += 1
                Goto, DGN
            }
        }
    }
    Until, looper=ArrayCount
}
Else
{
    TotalTime := (A_TickCount - startTime)/1000
    MsgBox, 0, done, 
    (LTrim
    DONE!
    %TotalTime% Seconds
    ), 10
    WinClose, construction.com - the construction industry marketplace - Internet Explorer
    Sleep, 333
}
Xl.Quit(), Xl := ""
Return


F8::ExitApp

F12::Pause

IELoad(Pwb)
{
	While !(Pwb.busy)
		Sleep, 100
	While (Pwb.busy)
		Sleep, 100
	While !(Pwb.document.Readystate = "Complete")
		Sleep, 100
}

/*
PMC File Version 5.0.5
---[Do not edit anything in this section]---

[PMC Code v5.0.5]|F3||1|Window,2,Fast,0,1,Input,-1,-1,1|1|Macro1
1|FileSelectFile|loginList, , C:\Users\rijul.kumar\Desktop\DGN Bots\Current Versions\DGN Planroom Log In.xlsx, Please select your Log In File|1|0|FileSelectFile|||||
2|Xl||1|0|COMInterface|Excel.Application||||
3|[Expression]|Xl.Workbooks.Open(loginList)|1|0|Expression|||||
4|[Expression]|endRow := Xl.columns("A").end(-4121).row|1|0|Expression|||||
5|[Assign Variable]|associations := %endRow% -1|1|0|Variable|Expression||||
6|[Assign Variable]|texty := |1|0|Variable|||||
7|[Assign Variable]|textyRow := 2|1|0|Variable|||||
8|[LoopStart]|textyRow <= endRow|1|0|While|||||
9|[Assign Variable]|name := Xl.range("A" . %textyRow%).value . " = " . (%textyRow%-1) . "`n`n"|1|0|Variable|Expression||||
10|[Concatenate Variable]|texty .= %name%|1|0|Variable|||||
11|[Add Variable]|textyRow += 1|1|0|Variable|||||
12|[LoopEnd]|LoopEnd|1|0|Loop|||||
13|InputBox|looper, Where would you like to start your syncing process?, %texty%, , 500, 700, , , , , 14|1|0|InputBox|||||
14|[Assign Variable]|startTime := %A_TickCount%|1|0|Variable|||||
15|Method:Navigate:|network2.construction.com|1|0|IECOM_Set|:|LoadWait|||
16|WinMaximize||1|333|WinMaximize||construction.com - the construction industry marketplace - Internet Explorer|||
17|WinActivate||1|333|WinActivate||construction.com - the construction industry marketplace - Internet Explorer|||
18|[Label]|DGN|1|0|Label|||||
19|Compare Variables|looper <= %associations%|1|0|If_Statement|looper=ArrayCount||||
20|[LoopStart]|LoopStart|0|0|Loop|looper=ArrayCount||||
21|[Assign Variable]|rowNum := %looper%+1|1|0|Variable|Expression||||
22|[Assign Variable]|userName := Xl.range("A" . %rowNum%).value|1|0|Variable|Expression||||
23|[Assign Variable]|passWord := Xl.range("B" . %rowNum%).value|1|0|Variable|Expression||||
24|[Assign Variable]|SSID := Xl.range("C" . %rowNum%).value|1|0|Variable|Expression||||
25|[Text]|%userName%|1|0|Send|||sign in||
26|[Text]|{Tab}|1|1|SendEvent|||sign in||
27|[Text]|%passWord%|1|0|Send|||sign in||
28|[Text]|{Enter}|1|1|SendEvent|||sign in||
29|[Pause]||1|500|Sleep|||||
30|Method:Navigate:|http://network2.construction.com/Home.aspx|1|0|IECOM_Set|:|LoadWait|||
31|Method:Navigate:|%SSID%|1|0|IECOM_Set|:|LoadWait|||
32|[Assign Variable]|first := 0|1|0|Variable|||||
33|Property:InnerHTML:ID|projBool|1|0|IECOM_Get|ctl00_contentPlaceHolderHeader_pcTop_listProjectCountText:||||
34|Compare Variables|projBool != |1|0|If_Statement|ctl00_contentPlaceHolderHeader_rptPager_lblTotalPageCount:||||
35|Method:Click:Name||1|0|IECOM_Set|project-select-all:0||||
36|Method:Click:Name||1|0|IECOM_Set|ctl00$contentPlaceHolderHeader$pcTop$HeaderActions$btnprjresltAction:0||||
37|Method:Click:ID||1|0|IECOM_Set|lnkViewProjects:||||
38|[Else]|Else|1|0|If_Statement|||||
39|[Add Variable]|looper += 1|1|0|Variable|||||
40|Method:Click:ID||1|0|IECOM_Set|ctl00_ucHeader_lnk_SignOut:|LoadWait|||
41|[Goto]|DGN|1|0|Goto|||||
42|[End If]|EndIf|1|0|If_Statement|||||
43|Continue, Continue, FoundX, FoundY|0, 0, %A_ScreenWidth%, %A_ScreenHeight%, 0x2DB98D, 0, Fast RGB|1|0|PixelSearch|UntilFound|Window|||
44|Property:InnerText:ID|projNum|1|0|IECOM_Get|ctl00_contentPlaceHolderHeader_rptPager_lblTotalPageCount:||||
45|Compare Variables|projNum != |1|0|If_Statement|ctl00_contentPlaceHolderHeader_rptPager_lblTotalPageCount:||||
46|[Assign Variable]|loopNum := %projNum%|1|0|Variable|||||
47|[Else]|Else|1|0|If_Statement|||||
48|[Assign Variable]|loopNum := 1|1|0|Variable|||||
49|[End If]|EndIf|1|0|If_Statement|||||
50|[Pause]||1|1000|Sleep|||||
51|[LoopStart]|loopNum>=0|1|0|While|||||
52|Compare Variables|loopNum > 0|1|0|If_Statement|UntilFound||||
53|[Label]|2|1|0|Label|||||
54|Method:Click:ClassName||1|0|IECOM_Set|syncRefreshText:0||||
55|Compare Variables|ErrorLevel != 0|1|0|If_Statement|syncRefreshText:0||||
56|[Goto]|2|1|0|Goto|||||
57|[End If]|EndIf|1|0|If_Statement|||||
58|[Label]|3|1|0|Label|||||
59|Method:Click:ClassName||1|0|IECOM_Set|planRoomOkText:1||||
60|Compare Variables|ErrorLevel != 0|1|0|If_Statement|syncRefreshText:0||||
61|[Goto]|3|1|0|Goto|||||
62|[End If]|EndIf|1|0|If_Statement|||||
63|[Label]|4|1|0|Label|||||
64|Method:Click:ID||1|0|IECOM_Set|lnkTrackProjects:||||
65|Compare Variables|ErrorLevel != 0|1|0|If_Statement|syncRefreshText:0||||
66|[Goto]|4|1|0|Goto|||||
67|[End If]|EndIf|1|0|If_Statement|||||
68|[Label]|5|1|0|Label|||||
69|Method:Click:ClassName||1|0|IECOM_Set|trackCheck:1||||
70|Compare Variables|ErrorLevel != 0|1|0|If_Statement|syncRefreshText:0||||
71|[Goto]|5|1|0|Goto|||||
72|[End If]|EndIf|1|0|If_Statement|||||
73|[Label]|6|1|0|Label|||||
74|Method:Click:ClassName||1|0|IECOM_Set|track-popup-submit:0||||
75|Compare Variables|ErrorLevel != 0|1|0|If_Statement|syncRefreshText:0||||
76|[Goto]|6|1|0|Goto|||||
77|[End If]|EndIf|1|0|If_Statement|||||
78|[Label]|7|1|0|Label|||||
79|Method:Focus:ClassName||1|0|IECOM_Set|syncRefreshText:0||||
80|Compare Variables|ErrorLevel != 0|1|0|If_Statement|syncRefreshText:0||||
81|[Goto]|7|1|0|Goto|||||
82|[Else]|Else|1|0|If_Statement|||||
83|Compare Variables|loopNum > 1|1|0|If_Statement|UntilFound||||
84|Method:Click:ID||1|0|IECOM_Set|ctl00_contentPlaceHolderHeader_rptPager_lblNext:|LoadWait|||
85|[End If]|EndIf|1|0|If_Statement|||||
86|[End If]|EndIf|1|0|If_Statement|||||
87|[Subtract Variable]|loopNum -= 1|1|0|Variable|||||
88|[Else]|Else|1|0|If_Statement|||||
89|Method:Click:ID||1|0|IECOM_Set|ctl00_ucHeader_lnk_SignOut:|LoadWait|||
90|[Pause]||1|500|Sleep|||||
91|[Add Variable]|looper += 1|1|0|Variable|||||
92|[Goto]|DGN|1|0|Goto|||||
93|[End If]|EndIf|1|0|If_Statement|||||
94|[LoopEnd]|LoopEnd|1|0|Loop|||||
95|[LoopEnd]|LoopEnd|1|0|Loop|||||
96|[Else]|Else|1|0|If_Statement|||||
97|[Assign Variable]|TotalTime := (A_TickCount - startTime)/1000|1|0|Variable|Expression||||
98|[MsgBox]|DONE!`n%TotalTime% Seconds|1|10|MsgBox|0|done|||
99|WinClose||1|333|WinClose||construction.com - the construction industry marketplace - Internet Explorer|||
100|[End If]|EndIf|1|0|If_Statement|||||
101|[Expression]|Xl.Quit(), Xl := ""|1|0|Expression|||||

*/
