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
password := ["welcome1", "welcome1", "welcome1", "welcome1", "exchange1175", "welcome1", "welcome1", "Welcome1", "Planroom2017", "welcome1", "welcome1", "planroom1", "welcome1"]
ArrayCount := userName.Length()
InputBox, loopNum, Where would you like to start your update process?, %texty%, , 500, 500, , , , , 1
If !IsObject(ie)
	ie := ComObjCreate("InternetExplorer.Application")
ie.Visible := true
ie.Navigate("planroom.construction.com")
IELoad(ie)
WinMaximize, Sign In - DODGE PlanRoom - Internet Explorer
Sleep, 333
Loop
{
    Loop
    {
        CoordMode, Pixel, Window
        ImageSearch, FoundX, FoundY, 0, 0, %A_ScreenWidth%, %A_ScreenHeight%, C:\Users\rijul.kumar\AppData\Roaming\MacroCreator\Screenshots\Screen_20170512141713.png
    }
    Until ErrorLevel = 0
    ie.document.getElementsByName("UserName")[0].Focus("")
    Send, % userName[loopNum]  ; sign in
    Send, {Tab}  ; sign in
    Send, % password[loopNum]  ; sign in
    Sleep, 200
    ie.document.getElementsByClassName("btn btn-default sign-in-button")[0].Click("")
    IELoad(ie)
    Loop
    {
        CoordMode, Pixel, Window
        PixelSearch, FoundX, FoundY, 0, 0, %A_ScreenWidth%, %A_ScreenHeight%, 0xDC6767, 0, Fast RGB
    }
    Until ErrorLevel = 0
    ie.Navigate("http://planroom.construction.com/3010100/planroom/my-projects/active?force=false")
    IELoad(ie)
    Loop
    {
        CoordMode, Pixel, Window
        PixelSearch, FoundX, FoundY, 0, 0, %A_ScreenWidth%, %A_ScreenHeight%, 0xDC6767, 0, Fast RGB
    }
    Until ErrorLevel = 0
    totalItems := ie.document.getElementsByClassName("ng-binding")[233].InnerText
    StringTrimLeft, totalItems, totalItems, 13
    ie.document.getElementsByClassName("ng-pristine ng-untouched ng-valid")[0].Focus("")
    SendRaw, bidding  ; sign in
    Sleep, 1000
    biddingItems := ie.document.getElementsByClassName("ng-binding")[233].InnerText
    StringTrimLeft, biddingItems, biddingItems, 13
    syncNum := totalItems-biddingItems
    ie.document.getElementsByClassName("input-group-addon glyphicon glyphicon-remove clear-icon")[0].Click("")
    Sleep, 500
    While syncNum>0
    {
        Loop
        {
            CoordMode, Pixel, Window
            PixelSearch, FoundX, FoundY, 0, 0, %A_ScreenWidth%, %A_ScreenHeight%, 0xDC6767, 0, Fast RGB
        }
        Until ErrorLevel = 0
        Loop
        {
            CoordMode, Pixel, Window
            ImageSearch, FoundX, FoundY, 0, 0, %A_ScreenWidth%, %A_ScreenHeight%, C:\Users\rijul.kumar\AppData\Roaming\MacroCreator\Screenshots\Screen_20170512113920.png
            CenterImgSrchCoords("C:\Users\rijul.kumar\AppData\Roaming\MacroCreator\Screenshots\Screen_20170512113920.png", FoundX, FoundY)
        }
        Until ErrorLevel = 0
        Click, %FoundX%, %FoundY% Left, 2
        Sleep, 100
        Sleep, 500
        ie.document.getElementsByClassName("ui-grid-cell-contents ng-binding ng-scope")[0].Click("")
        Sleep, 100
        ie.document.getElementsByClassName("ng-binding")[20].Click("")
        Click, Left, 1
        Sleep, 100
        WinMaximize, A
        Sleep, 333
        Loop
        {
            CoordMode, Pixel, Window
            PixelSearch, FoundX, FoundY, 0, 0, %A_ScreenWidth%, %A_ScreenHeight%, 0xDC6767, 0, Fast RGB
        }
        Until ErrorLevel = 0
        Loop
        {
            CoordMode, Pixel, Window
            ImageSearch, FoundX, FoundY, 0, 0, %A_ScreenWidth%, %A_ScreenHeight%, C:\Users\rijul.kumar\AppData\Roaming\MacroCreator\Screenshots\Screen_20170512130916.png
            CenterImgSrchCoords("C:\Users\rijul.kumar\AppData\Roaming\MacroCreator\Screenshots\Screen_20170512130916.png", FoundX, FoundY)
        }
        Until ErrorLevel = 0
        Click, %FoundX%, %FoundY% Left, 1
        Sleep, 100
        Loop
        {
            CoordMode, Pixel, Window
            ImageSearch, FoundX, FoundY, 0, 0, %A_ScreenWidth%, %A_ScreenHeight%, C:\Users\rijul.kumar\AppData\Roaming\MacroCreator\Screenshots\Screen_20170512131014.png
            CenterImgSrchCoords("C:\Users\rijul.kumar\AppData\Roaming\MacroCreator\Screenshots\Screen_20170512131014.png", FoundX, FoundY)
        }
        Until ErrorLevel = 0
        Click, %FoundX%, %FoundY% Left, 1
        Sleep, 100
        Loop
        {
            CoordMode, Pixel, Window
            ImageSearch, FoundX, FoundY, 0, 0, %A_ScreenWidth%, %A_ScreenHeight%, C:\Users\rijul.kumar\AppData\Roaming\MacroCreator\Screenshots\Screen_20170512133624.png
            CenterImgSrchCoords("C:\Users\rijul.kumar\AppData\Roaming\MacroCreator\Screenshots\Screen_20170512133624.png", FoundX, FoundY)
        }
        Until ErrorLevel = 0
        Loop, 10
        {
            Send, {Tab}
        }
        Sleep, 100
        Send, {Space}
        Sleep, 25
        Loop
        {
            CoordMode, Pixel, Window
            ImageSearch, FoundX, FoundY, 0, 0, %A_ScreenWidth%, %A_ScreenHeight%, C:\Users\rijul.kumar\AppData\Roaming\MacroCreator\Screenshots\Screen_20170512131257.png
            CenterImgSrchCoords("C:\Users\rijul.kumar\AppData\Roaming\MacroCreator\Screenshots\Screen_20170512131257.png", FoundX, FoundY)
        }
        Until ErrorLevel = 0
        Click, %FoundX%, %FoundY% Left, 1
        Sleep, 100
        Sleep, 500
        FoundX += 110
        FoundY -= 60
        Click, %FoundX%, %FoundY% Left, 1
        Sleep, 100
        Loop
        {
            CoordMode, Pixel, Window
            ImageSearch, FoundX, FoundY, 0, 0, %A_ScreenWidth%, %A_ScreenHeight%, C:\Users\rijul.kumar\AppData\Roaming\MacroCreator\Screenshots\Screen_20170512133624.png
            CenterImgSrchCoords("C:\Users\rijul.kumar\AppData\Roaming\MacroCreator\Screenshots\Screen_20170512133624.png", FoundX, FoundY)
        }
        Until ErrorLevel = 0
        Click, Left, 1
        Sleep, 100
        Loop, 31
        {
            Send, {Tab}
        }
        Sleep, 300
        Loop, 2
        {
            Send, {Down}
            Sleep, 10
        }
        Send, {Enter}
        Sleep, 25
        Loop, 3
        {
            Send, {Tab}
            Sleep, 5
        }
        Send, {Enter}
        Sleep, 25
        syncNum -= 1
        Click, Left, 1
        Sleep, 100
        Send, ^w
        Loop
        {
            CoordMode, Pixel, Window
            PixelSearch, FoundX, FoundY, 0, 0, %A_ScreenWidth%, %A_ScreenHeight%, 0xDC6767, 0, Fast RGB
        }
        Until ErrorLevel = 0
        Click, Left, 1
        Sleep, 100
        ie.Navigate("http://planroom.construction.com/3010100/planroom/my-projects/active?force=false")
        IELoad(ie)
        Loop
        {
            CoordMode, Pixel, Window
            PixelSearch, FoundX, FoundY, 0, 0, %A_ScreenWidth%, %A_ScreenHeight%, 0xDC6767, 0, Fast RGB
        }
        Until ErrorLevel = 0
    }
    ie.document.getElementsByTagName("A")[2].Click("")
    IELoad(ie)
    loopNum += 1  ; sign out
}
Until, ArrayCount<loopNum
/*
MsgBox, 48, Finito, DONE!
*/
Send, ^w
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

CenterImgSrchCoords(File, ByRef CoordX, ByRef CoordY)
{
	static LoadedPic
	LastEL := ErrorLevel
	Gui, Pict:Add, Pic, vLoadedPic, %File%
	GuiControlGet, LoadedPic, Pict:Pos
	Gui, Pict:Destroy
	CoordX += LoadedPicW // 2
	CoordY += LoadedPicH // 2
	ErrorLevel := LastEL
}

/*
PMC File Version 5.0.5
---[Do not edit anything in this section]---

[PMC Code v5.0.5]|F3||1|Window,2,Fast,0,1,Input,-1,-1,1|1|Macro1
1|[Assign Variable]|texty := shillebert@agcks.org = 1`n`nagcmo-admin@acmecontracting.net = 2`n`nkarly.hartford@agcok.com = 3`n`nkristys@lagc.org = 4`n`nSSleeper@bx.org = 5`n`nsuz@nesca.org = 6`n`ncea-admin@acmecontracting.net = 7`n`ntammy@ovabc.org = 8`n`nagc-ca-pradmin@acmecontracting.net = 9`n`nsta-ny--admin@acmecontracting.com = 10`n`nsmullane@agcwa.com = 11`n`nbrodgers@agc-utah.org = 12`n`nmgifford@agccolorado.org = 13|1|0|Variable|||||
2|[Assign Variable]|userName := ["shillebert@agcks.org", "agcmo-admin@acmecontracting.net", "karly.hartford@agcok.com", "kristys@lagc.org", "ssleeper@bx.org", "suz@nesca.org", "cea-admin@acmecontracting.net", "tammy@ovabc.org", "agc-ca-pradmin@acmecontracting.net", "sta-ny-admin@acmecontracting.com", "smullane@agcwa.com", "brodgers@agc-utah.org", "mgifford@agccolorado.org"]|1|0|Variable|Expression||||
3|[Assign Variable]|password := ["welcome1", "welcome1", "welcome1", "welcome1", "exchange1175", "welcome1", "welcome1", "Welcome1", "Planroom2017", "welcome1", "welcome1", "planroom1", "welcome1"]|1|0|Variable|Expression||||
4|[Assign Variable]|ArrayCount := userName.Length()|1|0|Variable|Expression||||
5|InputBox|loopNum, Where would you like to start your update process?, %texty%, , 500, 500, , , , , 1|1|0|InputBox|||||
6|Method:Navigate:|planroom.construction.com|1|0|IECOM_Set|:|LoadWait|||
7|WinMaximize||1|333|WinMaximize||Sign In - DODGE PlanRoom - Internet Explorer|||
8|[LoopStart]|LoopStart|0|0|Loop|ArrayCount<loopNum||||
9|Continue, Continue, FoundX, FoundY, 0|0, 0, %A_ScreenWidth%, %A_ScreenHeight%, C:\Users\rijul.kumar\AppData\Roaming\MacroCreator\Screenshots\Screen_20170512141713.png|1|0|ImageSearch|UntilFound|Window|||
10|Method:Focus:Name||1|0|IECOM_Set|UserName:0||||
11|[Text]|% userName[loopNum]|1|0|Send|||sign in||
12|[Text]|{Tab}|1|0|Send|||sign in||
13|[Text]|% password[loopNum]|1|200|Send|||sign in||
14|Method:Click:ClassName||1|0|IECOM_Set|btn btn-default sign-in-button:0|LoadWait|||
15|Continue, Continue, FoundX, FoundY|0, 0, %A_ScreenWidth%, %A_ScreenHeight%, 0xDC6767, 0, Fast RGB|1|0|PixelSearch|UntilFound|Window|||
16|Method:Navigate:|http://planroom.construction.com/3010100/planroom/my-projects/active?force=false|1|0|IECOM_Set|:|LoadWait|||
17|Continue, Continue, FoundX, FoundY|0, 0, %A_ScreenWidth%, %A_ScreenHeight%, 0xDC6767, 0, Fast RGB|1|0|PixelSearch|UntilFound|Window|||
18|Property:InnerText:ClassName|totalItems|1|0|IECOM_Get|ng-binding:233||||
19|StringTrimLeft|totalItems, totalItems, 13|1|0|StringTrimLeft|||||
20|Method:Focus:ClassName||1|0|IECOM_Set|ng-pristine ng-untouched ng-valid:0||||
21|[Text]|bidding|1|0|SendRaw|||sign in||
22|[Pause]||1|1000|Sleep|||||
23|Property:InnerText:ClassName|biddingItems|1|0|IECOM_Get|ng-binding:233||||
24|StringTrimLeft|biddingItems, biddingItems, 13|1|0|StringTrimLeft|||||
25|[Assign Variable]|syncNum := totalItems-biddingItems|1|0|Variable|Expression||||
26|Method:Click:ClassName||1|0|IECOM_Set|input-group-addon glyphicon glyphicon-remove clear-icon:0||||
27|[Pause]||1|500|Sleep|||||
28|[LoopStart]|syncNum>0|1|0|While|||||
29|Continue, Continue, FoundX, FoundY|0, 0, %A_ScreenWidth%, %A_ScreenHeight%, 0xDC6767, 0, Fast RGB|1|0|PixelSearch|UntilFound|Window|||
30|Continue, Continue, FoundX, FoundY, 1|0, 0, %A_ScreenWidth%, %A_ScreenHeight%, C:\Users\rijul.kumar\AppData\Roaming\MacroCreator\Screenshots\Screen_20170512113920.png|1|0|ImageSearch|UntilFound|Window|||
31|Left Move & Click|%FoundX%, %FoundY% Left, 2|1|100|Click|||||
32|[Pause]||1|500|Sleep|||||
33|Method:Click:ClassName||1|0|IECOM_Set|ui-grid-cell-contents ng-binding ng-scope:0||||
34|[Pause]||1|100|Sleep|||||
35|Method:Click:ClassName||1|0|IECOM_Set|ng-binding:20||||
36|Left Click|Left, 1, |1|100|Click|||||
37|WinMaximize||1|333|WinMaximize||A|||
38|Continue, Continue, FoundX, FoundY|0, 0, %A_ScreenWidth%, %A_ScreenHeight%, 0xDC6767, 0, Fast RGB|1|0|PixelSearch|UntilFound|Window|||
39|Continue, Continue, FoundX, FoundY, 1|0, 0, %A_ScreenWidth%, %A_ScreenHeight%, C:\Users\rijul.kumar\AppData\Roaming\MacroCreator\Screenshots\Screen_20170512130916.png|1|0|ImageSearch|UntilFound|Window|||
40|Left Move & Click|%FoundX%, %FoundY% Left, 1|1|100|Click|||||
41|Continue, Continue, FoundX, FoundY, 1|0, 0, %A_ScreenWidth%, %A_ScreenHeight%, C:\Users\rijul.kumar\AppData\Roaming\MacroCreator\Screenshots\Screen_20170512131014.png|1|0|ImageSearch|UntilFound|Window|||
42|Left Move & Click|%FoundX%, %FoundY% Left, 1|1|100|Click|||||
43|Continue, Continue, FoundX, FoundY, 1|0, 0, %A_ScreenWidth%, %A_ScreenHeight%, C:\Users\rijul.kumar\AppData\Roaming\MacroCreator\Screenshots\Screen_20170512133624.png|1|0|ImageSearch|UntilFound|Window|||
44|[Text]|{Tab}|10|0|Send|||||
45|[Pause]||1|100|Sleep|||||
46|[Text]|{Space}|1|25|Send|||||
47|Continue, Continue, FoundX, FoundY, 1|0, 0, %A_ScreenWidth%, %A_ScreenHeight%, C:\Users\rijul.kumar\AppData\Roaming\MacroCreator\Screenshots\Screen_20170512131257.png|1|0|ImageSearch|UntilFound|Window|||
48|Left Move & Click|%FoundX%, %FoundY% Left, 1|1|100|Click|||||
49|[Pause]||1|500|Sleep|||||
50|[Add Variable]|FoundX += 110|1|0|Variable|||||
51|[Subtract Variable]|FoundY -= 60|1|0|Variable|||||
52|Left Move & Click|%FoundX%, %FoundY% Left, 1|1|100|Click|||||
53|Continue, Continue, FoundX, FoundY, 1|0, 0, %A_ScreenWidth%, %A_ScreenHeight%, C:\Users\rijul.kumar\AppData\Roaming\MacroCreator\Screenshots\Screen_20170512133624.png|1|0|ImageSearch|UntilFound|Window|||
54|Left Click|Left, 1, |1|100|Click|||||
55|[Text]|{Tab}|31|0|Send|||||
56|[Pause]||1|300|Sleep|||||
57|[Text]|{Down}|2|10|Send|||||
58|[Text]|{Enter}|1|25|Send|||||
59|[Text]|{Tab}|3|5|Send|||||
60|[Text]|{Enter}|1|25|Send|||||
61|[Subtract Variable]|syncNum -= 1|1|0|Variable|||||
62|Left Click|Left, 1, |1|100|Click|||||
63|[Text]|^w|1|0|Send|||||
64|Continue, Continue, FoundX, FoundY|0, 0, %A_ScreenWidth%, %A_ScreenHeight%, 0xDC6767, 0, Fast RGB|1|0|PixelSearch|UntilFound|Window|||
65|Left Click|Left, 1, |1|100|Click|||||
66|Method:Navigate:|http://planroom.construction.com/3010100/planroom/my-projects/active?force=false|1|0|IECOM_Set|:0|LoadWait|||
67|Continue, Continue, FoundX, FoundY|0, 0, %A_ScreenWidth%, %A_ScreenHeight%, 0xDC6767, 0, Fast RGB|1|0|PixelSearch|UntilFound|Window|||
68|[LoopEnd]|LoopEnd|1|0|Loop|||||
69|Method:Click:TagName||1|0|IECOM_Set|A:2|LoadWait|||
70|[Add Variable]|loopNum += 1|1|0|Variable|||sign out||
71|[LoopEnd]|LoopEnd|1|0|Loop|||||
072|[MsgBox]|DONE!|1|0|MsgBox|48|Finito|||
73|[Text]|^w|1|0|Send|||||

*/
