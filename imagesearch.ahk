#Singleinstance, Force

;esc::
    ;ExitApp
    ;return

MyFilePath := "C:\Users\Ronald\Desktop\AUtoHotKeyTuts\names1.xlsx"

Try oExcel := ComObjActive("Excel.Application")
Try oWorkBook := oExcel.Workbooks.Open(MyFilePath)
if(!oWorkBook)
{
    oExcel := ComObjCreate("Excel.Application")
    oWorkBook := oExcel.Workbooks.Open(MyFilePath)
}
;oExcel .visible := False
sleep 500
WinMinimize, names1 - Excel ahk_class XLMAIN
sleep 500
mysheet := oworkbook.sheets("Sheet1")
mycells := mysheet.range("A1:A5")

CoordMode Pixel, screen
CoordMode mouse, screen

imagefile = rocky.png

ImageSearch X, Y, 0, 0, A_ScreenWidth, A_ScreenHeight,*32 *w0 *h0 %imagefile%

sleep 500

gui,add,picture,hwndmypic,%imagefile%
controlgetpos,,,width, height,,ahk_id %mypic%

winactivate, Nice4What - Excel

For cell in mycells
{
    Mousemove, X+(width//2), Y+(height//2), 10
    mousemove, -310,-10,10,R
    Click
    sleep 500
    ;Click
    Send {Delete}
    sleep 500
    cell1 := cell.value
    send %cell1%
    sleep 500
    Send, {Enter}
    
    cell2 := cell.offset(0,1).value
    mousemove, 50,0,10,R
    Click
    sleep 500
    ;Click
    Send {Delete}
    sleep 500
    send, %cell2%
    
    sleep 500
    Send, {Enter}
    sleep 500
    cell.offset(0,2) := "Done"
}
/*if ErrorLevel = 2
    MsgBox Could not conduct the search.
else if ErrorLevel = 1
    MsgBox Icon could not be found on the screen.
else
    MsgBox The icon was found at %X%x%Y%.
    
Gui, add, picture, , orange.png
Gui, show
*/