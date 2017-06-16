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
texty := 
(LTrim
"shillebert@agcks.org = 1

agcmo-admin@acmecontracting.net = 2

karly.hartford@agcok.com = 3

kristys@lagc.org = 4

SSleeper@bx.org = 5

suz@nesca.org = 6

cea-admin@acmecontracting.net = 7

tammy@ovabc.org = 8

agc-ca-pradmin@acmecontracting.net = 9

sta-ny--admin@acmecontracting.com = 10

smullane@agcwa.com = 11

brodgers@agc-utah.org = 12

mgifford@agccolorado.org = 13"
)
userName := ["shillebert@agcks.org", "agcmo-admin@acmecontracting.net", "karly.hartford@agcok.com", "kristys@lagc.org", "ssleeper@bx.org", "suz@nesca.org", "cea-admin@acmecontracting.net", "tammy@ovabc.org", "agc-ca-pradmin@acmecontracting.net", "sta-ny-admin@acmecontracting.com", "smullane@agcwa.com", "brodgers@agc-utah.org", "mgifford@agccolorado.org"]
password := ["welcome1", "welcome1", "welcome1", "Welcome1", "exchange1175", "welcome1", "welcome1", "Welcome1", "Planroom2017", "welcome1", "welcome1", "planroom1", "welcome1"]
ssid := ["http://network2.construction.com/ProjectResults.aspx?ssid=11329215" ,"http://network2.construction.com/ProjectResults.aspx?ssid=11329216" ,"http://network2.construction.com/ProjectResults.aspx?ssid=11329217" ,"http://network2.construction.com/ProjectResults.aspx?ssid=11329220" ,"http://network2.construction.com/ProjectResults.aspx?ssid=11329246" ,"http://network2.construction.com/ProjectResults.aspx?ssid=11329248" ,"http://network2.construction.com/ProjectResults.aspx?ssid=11329252" ,"http://network2.construction.com/ProjectResults.aspx?ssid=11329259" ,"http://network2.construction.com/ProjectResults.aspx?ssid=11329260" ,"http://network2.construction.com/ProjectResults.aspx?ssid=11329264" ,"http://network2.construction.com/ProjectResults.aspx?ssid=11329268" ,"http://network2.construction.com/ProjectResults.aspx?ssid=11329269" ,"http://network2.construction.com/ProjectResults.aspx?ssid=11329271"]
ArrayCount := userName.Length()
InputBox, looper, Where would you like to start your update process?, %texty%, , 500, 700, , , , , 1
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
If looper <= %ArrayCount%
{
    Loop
    {
        Send, % userName[looper]  ; sign in
        CurrentKeyDelay := A_KeyDelay
        SetKeyDelay, 1
        SendEvent, {Tab}  ; sign in
        SetKeyDelay, %CurrentKeyDelay%
        Send, % password[looper]  ; sign in
        CurrentKeyDelay := A_KeyDelay
        SetKeyDelay, 1
        SendEvent, {Enter}  ; sign in
        SetKeyDelay, %CurrentKeyDelay%
        Sleep, 500
        ie.Navigate("http://network2.construction.com/Home.aspx")
        IELoad(ie)
        ie.Navigate(ssid[looper])
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
                Loop
                {
                    CoordMode, Pixel, Window
                    PixelSearch, FoundX, FoundY, 0, 0, %A_ScreenWidth%, %A_ScreenHeight%, 0x2DB98D, 0, Fast RGB
                }
                Until ErrorLevel = 0
                If ErrorLevel = 0
                {
                    ie.document.getElementsByClassName("syncRefreshText")[0].Click("")
                }
                Loop
                {
                    CoordMode, Pixel, Window
                    PixelSearch, FoundX, FoundY, 0, 0, %A_ScreenWidth%, %A_ScreenHeight%, 0x61CAA9, 0, Fast RGB
                }
                Until ErrorLevel = 0
                If ErrorLevel = 0
                {
                    ie.document.getElementsByClassName("planRoomOkText")[1].Click("")
                }
                Loop
                {
                    CoordMode, Pixel, Window
                    PixelSearch, FoundX, FoundY, 0, 0, %A_ScreenWidth%, %A_ScreenHeight%, 0x2DB98D, 0, Fast RGB
                }
                Until ErrorLevel = 0
                If ErrorLevel = 0
                {
                    ie.document.getElementByID("lnkTrackProjects").Click("") := ""
                }
                Loop
                {
                    CoordMode, Pixel, Window
                    PixelSearch, FoundX, FoundY, 0, 0, %A_ScreenWidth%, %A_ScreenHeight%, 0x102C42, 0, Fast RGB
                }
                Until ErrorLevel = 0
                If ErrorLevel = 0
                {
                    Sleep, 300
                    ie.document.getElementsByClassName("trackCheck")[1].Click("")
                }
                Loop
                {
                    CoordMode, Pixel, Window
                    PixelSearch, FoundX, FoundY, 0, 0, %A_ScreenWidth%, %A_ScreenHeight%, 0x102C42, 0, Fast RGB
                }
                Until ErrorLevel = 0
                If ErrorLevel = 0
                {
                    Sleep, 300
                    ie.document.getElementsByClassName("track-popup-submit")[0].Click("")
                }
                Loop
                {
                    CoordMode, Pixel, Window
                    PixelSearch, FoundX, FoundY, 0, 0, %A_ScreenWidth%, %A_ScreenHeight%, 0x2DB98D, 0, Fast RGB
                }
                Until ErrorLevel = 0
                If ErrorLevel = 0
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
1|[Assign Variable]|texty := shillebert@agcks.org = 1`n`nagcmo-admin@acmecontracting.net = 2`n`nkarly.hartford@agcok.com = 3`n`nkristys@lagc.org = 4`n`nSSleeper@bx.org = 5`n`nsuz@nesca.org = 6`n`ncea-admin@acmecontracting.net = 7`n`ntammy@ovabc.org = 8`n`nagc-ca-pradmin@acmecontracting.net = 9`n`nsta-ny--admin@acmecontracting.com = 10`n`nsmullane@agcwa.com = 11`n`nbrodgers@agc-utah.org = 12`n`nmgifford@agccolorado.org = 13|1|0|Variable|||||
2|[Assign Variable]|userName := ["shillebert@agcks.org", "agcmo-admin@acmecontracting.net", "karly.hartford@agcok.com", "kristys@lagc.org", "ssleeper@bx.org", "suz@nesca.org", "cea-admin@acmecontracting.net", "tammy@ovabc.org", "agc-ca-pradmin@acmecontracting.net", "sta-ny-admin@acmecontracting.com", "smullane@agcwa.com", "brodgers@agc-utah.org", "mgifford@agccolorado.org"]|1|0|Variable|Expression||||
3|[Assign Variable]|password := ["welcome1", "welcome1", "welcome1", "Welcome1", "exchange1175", "welcome1", "welcome1", "Welcome1", "Planroom2017", "welcome1", "welcome1", "planroom1", "welcome1"]|1|0|Variable|Expression||||
4|[Assign Variable]|ssid := ["http://network2.construction.com/ProjectResults.aspx?ssid=11329215" ,"http://network2.construction.com/ProjectResults.aspx?ssid=11329216" ,"http://network2.construction.com/ProjectResults.aspx?ssid=11329217" ,"http://network2.construction.com/ProjectResults.aspx?ssid=11329220" ,"http://network2.construction.com/ProjectResults.aspx?ssid=11329246" ,"http://network2.construction.com/ProjectResults.aspx?ssid=11329248" ,"http://network2.construction.com/ProjectResults.aspx?ssid=11329252" ,"http://network2.construction.com/ProjectResults.aspx?ssid=11329259" ,"http://network2.construction.com/ProjectResults.aspx?ssid=11329260" ,"http://network2.construction.com/ProjectResults.aspx?ssid=11329264" ,"http://network2.construction.com/ProjectResults.aspx?ssid=11329268" ,"http://network2.construction.com/ProjectResults.aspx?ssid=11329269" ,"http://network2.construction.com/ProjectResults.aspx?ssid=11329271"]|1|0|Variable|Expression||||
5|[Assign Variable]|ArrayCount := userName.Length()|1|0|Variable|Expression||||
6|InputBox|looper, Where would you like to start your update process?, %texty%, , 500, 700, , , , , 1|1|0|InputBox|||||
7|[Assign Variable]|startTime := %A_TickCount%|1|0|Variable|||||
8|Method:Navigate:|network2.construction.com|1|0|IECOM_Set|:|LoadWait|||
9|WinMaximize||1|333|WinMaximize||construction.com - the construction industry marketplace - Internet Explorer|||
10|WinActivate||1|333|WinActivate||construction.com - the construction industry marketplace - Internet Explorer|||
11|[Label]|DGN|1|0|Label|||||
12|Compare Variables|looper <= %ArrayCount%|1|0|If_Statement|looper=ArrayCount||||
13|[LoopStart]|LoopStart|0|0|Loop|looper=ArrayCount||||
14|[Text]|% userName[looper]|1|0|Send|||sign in||
15|[Text]|{Tab}|1|1|SendEvent|||sign in||
16|[Text]|% password[looper]|1|0|Send|||sign in||
17|[Text]|{Enter}|1|1|SendEvent|||sign in||
18|[Pause]||1|500|Sleep|||||
19|Method:Navigate:|http://network2.construction.com/Home.aspx|1|0|IECOM_Set|:|LoadWait|||
20|Method:Navigate:|% ssid[looper]|1|0|IECOM_Set|:|LoadWait|||
21|[Assign Variable]|first := 0|1|0|Variable|||||
22|Property:InnerHTML:ID|projBool|1|0|IECOM_Get|ctl00_contentPlaceHolderHeader_pcTop_listProjectCountText:||||
23|Compare Variables|projBool != |1|0|If_Statement|ctl00_contentPlaceHolderHeader_rptPager_lblTotalPageCount:||||
24|Method:Click:Name||1|0|IECOM_Set|project-select-all:0||||
25|Method:Click:Name||1|0|IECOM_Set|ctl00$contentPlaceHolderHeader$pcTop$HeaderActions$btnprjresltAction:0||||
26|Method:Click:ID||1|0|IECOM_Set|lnkViewProjects:||||
27|[Else]|Else|1|0|If_Statement|||||
28|[Add Variable]|looper += 1|1|0|Variable|||||
29|Method:Click:ID||1|0|IECOM_Set|ctl00_ucHeader_lnk_SignOut:|LoadWait|||
30|[Goto]|DGN|1|0|Goto|||||
31|[End If]|EndIf|1|0|If_Statement|||||
32|Continue, Continue, FoundX, FoundY|0, 0, %A_ScreenWidth%, %A_ScreenHeight%, 0x2DB98D, 0, Fast RGB|1|0|PixelSearch|UntilFound|Window|||
33|Property:InnerText:ID|projNum|1|0|IECOM_Get|ctl00_contentPlaceHolderHeader_rptPager_lblTotalPageCount:||||
34|Compare Variables|projNum != |1|0|If_Statement|ctl00_contentPlaceHolderHeader_rptPager_lblTotalPageCount:||||
35|[Assign Variable]|loopNum := %projNum%|1|0|Variable|||||
36|[Else]|Else|1|0|If_Statement|||||
37|[Assign Variable]|loopNum := 1|1|0|Variable|||||
38|[End If]|EndIf|1|0|If_Statement|||||
39|[Pause]||1|1000|Sleep|||||
40|[LoopStart]|loopNum>=0|1|0|While|||||
41|Compare Variables|loopNum > 0|1|0|If_Statement|UntilFound||||
42|Continue, Continue, FoundX, FoundY|0, 0, %A_ScreenWidth%, %A_ScreenHeight%, 0x2DB98D, 0, Fast RGB|1|0|PixelSearch|UntilFound|Window|||
43|If Image/Pixel Found||1|0|If_Statement|||||
44|Method:Click:ClassName||1|0|IECOM_Set|syncRefreshText:0||||
45|[End If]|EndIf|1|0|If_Statement|||||
46|Continue, Continue, FoundX, FoundY|0, 0, %A_ScreenWidth%, %A_ScreenHeight%, 0x61CAA9, 0, Fast RGB|1|0|PixelSearch|UntilFound|Window|||
47|If Image/Pixel Found||1|0|If_Statement|||||
48|Method:Click:ClassName||1|0|IECOM_Set|planRoomOkText:1||||
49|[End If]|EndIf|1|0|If_Statement|||||
50|Continue, Continue, FoundX, FoundY|0, 0, %A_ScreenWidth%, %A_ScreenHeight%, 0x2DB98D, 0, Fast RGB|1|0|PixelSearch|UntilFound|Window|||
51|If Image/Pixel Found||1|0|If_Statement|||||
52|Method:Click:ID||1|0|IECOM_Set|lnkTrackProjects:||||
53|[End If]|EndIf|1|0|If_Statement|||||
54|Continue, Continue, FoundX, FoundY|0, 0, %A_ScreenWidth%, %A_ScreenHeight%, 0x102C42, 0, Fast RGB|1|0|PixelSearch|UntilFound|Window|||
55|If Image/Pixel Found||1|0|If_Statement|||||
56|[Pause]||1|300|Sleep|||||
57|Method:Click:ClassName||1|0|IECOM_Set|trackCheck:1||||
58|[End If]|EndIf|1|0|If_Statement|||||
59|Continue, Continue, FoundX, FoundY|0, 0, %A_ScreenWidth%, %A_ScreenHeight%, 0x102C42, 0, Fast RGB|1|0|PixelSearch|UntilFound|Window|||
60|If Image/Pixel Found||1|0|If_Statement|||||
61|[Pause]||1|300|Sleep|||||
62|Method:Click:ClassName||1|0|IECOM_Set|track-popup-submit:0||||
63|[End If]|EndIf|1|0|If_Statement|||||
64|Continue, Continue, FoundX, FoundY|0, 0, %A_ScreenWidth%, %A_ScreenHeight%, 0x2DB98D, 0, Fast RGB|1|0|PixelSearch|UntilFound|Window|||
65|If Image/Pixel Found||1|0|If_Statement|||||
66|Compare Variables|loopNum > 1|1|0|If_Statement|UntilFound||||
67|Method:Click:ID||1|0|IECOM_Set|ctl00_contentPlaceHolderHeader_rptPager_lblNext:|LoadWait|||
68|[End If]|EndIf|1|0|If_Statement|||||
69|[End If]|EndIf|1|0|If_Statement|||||
70|[Subtract Variable]|loopNum -= 1|1|0|Variable|||||
71|[Else]|Else|1|0|If_Statement|||||
72|Method:Click:ID||1|0|IECOM_Set|ctl00_ucHeader_lnk_SignOut:|LoadWait|||
73|[Pause]||1|500|Sleep|||||
74|[Add Variable]|looper += 1|1|0|Variable|||||
75|[Goto]|DGN|1|0|Goto|||||
76|[End If]|EndIf|1|0|If_Statement|||||
77|[LoopEnd]|LoopEnd|1|0|Loop|||||
78|[LoopEnd]|LoopEnd|1|0|Loop|||||
79|[Else]|Else|1|0|If_Statement|||||
80|[Assign Variable]|TotalTime := (A_TickCount - startTime)/1000|1|0|Variable|Expression||||
81|[MsgBox]|DONE!`n%TotalTime% Seconds|1|10|MsgBox|0|done|||
82|WinClose||1|333|WinClose||construction.com - the construction industry marketplace - Internet Explorer|||
83|[End If]|EndIf|1|0|If_Statement|||||

*/
