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
FileSelectFile, loginList, , C:\Users\rijul.kumar\Desktop\DGN Bots\Current Versions\DGN Planroom Log In.xlsx, Please select your Log In File, (*.csv)
UNarray := []
PWarray := []
SSIDarray := []
texty := ""
Loop, Read, %loginList%
{
    loopy := 1
    LineNumber := A_Index
    Loop, Parse, A_LoopReadLine, CSV
    {
        If LineNumber = 1
        {
            UNarray.Insert(A_LoopField)
            texty .= UNArray[A_Index]
            texty .= "="
            texty .= A_Index
            texty .= 
            (LTrim
            "
            
            "
            )
        }
        Else If LineNumber = 2
        {
            PWarray.Insert(A_LoopField)
        }
        Else If LineNumber = 3
        {
            SSIDarray.Insert(A_LoopField)
        }
        loopy += 1
    }
}
textyRow := 2
associations := % UNarray.Length()
InputBox, looper, Where would you like to start your syncing process?, %texty%, , 500, 500, , , , , 1
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
        userName := UNarray[looper]
        passWord := PWarray[looper]
        SSID := SSIDarray[looper]
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
1|FileSelectFile|loginList, , C:\Users\rijul.kumar\Desktop\DGN Bots\Current Versions\DGN Planroom Log In.xlsx, Please select your Log In File, (*.csv)|1|0|FileSelectFile|||||
2|[Assign Variable]|UNarray := []|1|0|Variable|Expression||||
3|[Assign Variable]|PWarray := []|1|0|Variable|Expression||||
4|[Assign Variable]|SSIDarray := []|1|0|Variable|Expression||||
5|[Assign Variable]|texty := |1|0|Variable|||||
6|[LoopStart]|%loginList%|1|0|LoopRead|||||
7|[Assign Variable]|loopy := 1|1|0|Variable|||||
8|[Assign Variable]|LineNumber := %A_Index%|1|0|Variable|||||
9|[LoopStart]|A_LoopReadLine`, CSV`, |1|0|LoopParse|||||
10|Compare Variables|LineNumber = 1|1|0|If_Statement|||||
11|[Expression]|UNarray.Insert(A_LoopField)|1|0|Expression|||||
12|[Concatenate Variable]|texty .= % UNArray[A_Index]|1|0|Variable|||||
13|[Concatenate Variable]|texty .= =|1|0|Variable|||||
14|[Concatenate Variable]|texty .= %A_Index%|1|0|Variable|||||
15|[Concatenate Variable]|texty .= `n`n|1|0|Variable|||||
16|[ElseIf] Compare Variables|LineNumber = 2|1|0|If_Statement|||||
17|[Expression]|PWarray.Insert(A_LoopField)|1|0|Expression|||||
18|[ElseIf] Compare Variables|LineNumber = 3|1|0|If_Statement|||||
19|[Expression]|SSIDarray.Insert(A_LoopField)|1|0|Expression|||||
20|[End If]|EndIf|1|0|If_Statement|||||
21|[Add Variable]|loopy += 1|1|0|Variable|||||
22|[LoopEnd]|LoopEnd|1|0|Loop|||||
23|[LoopEnd]|LoopEnd|1|0|Loop|||||
24|[Assign Variable]|textyRow := 2|1|0|Variable|||||
25|[Assign Variable]|associations := % UNarray.Length()|1|0|Variable|Expression||||
26|InputBox|looper, Where would you like to start your syncing process?, %texty%, , 500, 500, , , , , 1|1|0|InputBox|||||
27|[Assign Variable]|startTime := %A_TickCount%|1|0|Variable|||||
28|Method:Navigate:|network2.construction.com|1|0|IECOM_Set|:|LoadWait|||
29|WinMaximize||1|333|WinMaximize||construction.com - the construction industry marketplace - Internet Explorer|||
30|WinActivate||1|333|WinActivate||construction.com - the construction industry marketplace - Internet Explorer|||
31|[Label]|DGN|1|0|Label|||||
32|Compare Variables|looper <= %associations%|1|0|If_Statement|looper=ArrayCount||||
33|[LoopStart]|LoopStart|0|0|Loop|looper=ArrayCount||||
34|[Assign Variable]|userName := UNarray[looper]|1|0|Variable|Expression||||
35|[Assign Variable]|passWord := PWarray[looper]|1|0|Variable|Expression||||
36|[Assign Variable]|SSID := SSIDarray[looper]|1|0|Variable|Expression||||
37|[Text]|%userName%|1|0|Send|||sign in||
38|[Text]|{Tab}|1|1|SendEvent|||sign in||
39|[Text]|%passWord%|1|0|Send|||sign in||
40|[Text]|{Enter}|1|1|SendEvent|||sign in||
41|[Pause]||1|500|Sleep|||||
42|Method:Navigate:|http://network2.construction.com/Home.aspx|1|0|IECOM_Set|:|LoadWait|||
43|Method:Navigate:|%SSID%|1|0|IECOM_Set|:|LoadWait|||
44|[Assign Variable]|first := 0|1|0|Variable|||||
45|Property:InnerHTML:ID|projBool|1|0|IECOM_Get|ctl00_contentPlaceHolderHeader_pcTop_listProjectCountText:||||
46|Compare Variables|projBool != |1|0|If_Statement|ctl00_contentPlaceHolderHeader_rptPager_lblTotalPageCount:||||
47|Method:Click:Name||1|0|IECOM_Set|project-select-all:0||||
48|Method:Click:Name||1|0|IECOM_Set|ctl00$contentPlaceHolderHeader$pcTop$HeaderActions$btnprjresltAction:0||||
49|Method:Click:ID||1|0|IECOM_Set|lnkViewProjects:||||
50|[Else]|Else|1|0|If_Statement|||||
51|[Add Variable]|looper += 1|1|0|Variable|||||
52|Method:Click:ID||1|0|IECOM_Set|ctl00_ucHeader_lnk_SignOut:|LoadWait|||
53|[Goto]|DGN|1|0|Goto|||||
54|[End If]|EndIf|1|0|If_Statement|||||
55|Continue, Continue, FoundX, FoundY|0, 0, %A_ScreenWidth%, %A_ScreenHeight%, 0x2DB98D, 0, Fast RGB|1|0|PixelSearch|UntilFound|Window|||
56|Property:InnerText:ID|projNum|1|0|IECOM_Get|ctl00_contentPlaceHolderHeader_rptPager_lblTotalPageCount:||||
57|Compare Variables|projNum != |1|0|If_Statement|ctl00_contentPlaceHolderHeader_rptPager_lblTotalPageCount:||||
58|[Assign Variable]|loopNum := %projNum%|1|0|Variable|||||
59|[Else]|Else|1|0|If_Statement|||||
60|[Assign Variable]|loopNum := 1|1|0|Variable|||||
61|[End If]|EndIf|1|0|If_Statement|||||
62|[Pause]||1|1000|Sleep|||||
63|[LoopStart]|loopNum>=0|1|0|While|||||
64|Compare Variables|loopNum > 0|1|0|If_Statement|UntilFound||||
65|[Label]|2|1|0|Label|||||
66|Method:Click:ClassName||1|0|IECOM_Set|syncRefreshText:0||||
67|Compare Variables|ErrorLevel != 0|1|0|If_Statement|syncRefreshText:0||||
68|[Goto]|2|1|0|Goto|||||
69|[End If]|EndIf|1|0|If_Statement|||||
70|[Label]|3|1|0|Label|||||
71|Method:Click:ClassName||1|0|IECOM_Set|planRoomOkText:1||||
72|Compare Variables|ErrorLevel != 0|1|0|If_Statement|syncRefreshText:0||||
73|[Goto]|3|1|0|Goto|||||
74|[End If]|EndIf|1|0|If_Statement|||||
75|[Label]|4|1|0|Label|||||
76|Method:Click:ID||1|0|IECOM_Set|lnkTrackProjects:||||
77|Compare Variables|ErrorLevel != 0|1|0|If_Statement|syncRefreshText:0||||
78|[Goto]|4|1|0|Goto|||||
79|[End If]|EndIf|1|0|If_Statement|||||
80|[Label]|5|1|0|Label|||||
81|Method:Click:ClassName||1|0|IECOM_Set|trackCheck:1||||
82|Compare Variables|ErrorLevel != 0|1|0|If_Statement|syncRefreshText:0||||
83|[Goto]|5|1|0|Goto|||||
84|[End If]|EndIf|1|0|If_Statement|||||
85|[Label]|6|1|0|Label|||||
86|Method:Click:ClassName||1|0|IECOM_Set|track-popup-submit:0||||
87|Compare Variables|ErrorLevel != 0|1|0|If_Statement|syncRefreshText:0||||
88|[Goto]|6|1|0|Goto|||||
89|[End If]|EndIf|1|0|If_Statement|||||
90|[Label]|7|1|0|Label|||||
91|Method:Focus:ClassName||1|0|IECOM_Set|syncRefreshText:0||||
92|Compare Variables|ErrorLevel != 0|1|0|If_Statement|syncRefreshText:0||||
93|[Goto]|7|1|0|Goto|||||
94|[Else]|Else|1|0|If_Statement|||||
95|Compare Variables|loopNum > 1|1|0|If_Statement|UntilFound||||
96|Method:Click:ID||1|0|IECOM_Set|ctl00_contentPlaceHolderHeader_rptPager_lblNext:|LoadWait|||
97|[End If]|EndIf|1|0|If_Statement|||||
98|[End If]|EndIf|1|0|If_Statement|||||
99|[Subtract Variable]|loopNum -= 1|1|0|Variable|||||
100|[Else]|Else|1|0|If_Statement|||||
101|Method:Click:ID||1|0|IECOM_Set|ctl00_ucHeader_lnk_SignOut:|LoadWait|||
102|[Pause]||1|500|Sleep|||||
103|[Add Variable]|looper += 1|1|0|Variable|||||
104|[Goto]|DGN|1|0|Goto|||||
105|[End If]|EndIf|1|0|If_Statement|||||
106|[LoopEnd]|LoopEnd|1|0|Loop|||||
107|[LoopEnd]|LoopEnd|1|0|Loop|||||
108|[Else]|Else|1|0|If_Statement|||||
109|[Assign Variable]|TotalTime := (A_TickCount - startTime)/1000|1|0|Variable|Expression||||
110|[MsgBox]|DONE!`n%TotalTime% Seconds|1|10|MsgBox|0|done|||
111|WinClose||1|333|WinClose||construction.com - the construction industry marketplace - Internet Explorer|||
112|[End If]|EndIf|1|0|If_Statement|||||

*/
