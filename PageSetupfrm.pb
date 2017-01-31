;**********************************************************************
;
;    Page Setup Module For PB 5.4
;
;    FileName PageSetupfrm.pb
;
;******************************************************************
DeclareModule PSetup
  
  Declare Open()
  
EndDeclareModule

Module PSetup

  Global btn_Ok,frm_Margins, str_Top, str_Left, str_Bottom, str_Right.i

  Global TopMargin.i,LeftMargin.i,BottomMargin.i,RightMargin.i

Procedure.l Open()
  wnd_PageSetup = OpenWindow(#PB_Any, x, y, 540, 260, "", #PB_Window_TitleBar | #PB_Window_Tool | #PB_Window_ScreenCentered)
  btn_Cancel = ButtonGadget(#PB_Any, 430, 220, 100, 30, "Cancel")
  btn_Ok = ButtonGadget(#PB_Any, 320, 220, 100, 30, "Ok")
  FrameGadget(#PB_Any, 10, 110, 310, 100, " Margins")
  TextGadget(#PB_Any, 130, 140, 30, 20, "mm")
  TextGadget(#PB_Any, 130, 170, 30, 20, "mm")
  TextGadget(#PB_Any, 170, 140, 50, 20, "Left", #PB_Text_Right)
  TextGadget(#PB_Any, 20, 140, 50, 20, "Top", #PB_Text_Right)
  TextGadget(#PB_Any, 280, 140, 30, 20, "mm")
  TextGadget(#PB_Any, 280, 170, 30, 20, "mm")
  TextGadget(#PB_Any, 20, 170, 50, 20, "Bottom", #PB_Text_Right)
  TextGadget(#PB_Any, 170, 170, 50, 20, "Right", #PB_Text_Right)
  str_Top = StringGadget(#PB_Any, 80, 140, 40, 20, "")
  str_Bottom = StringGadget(#PB_Any, 80, 170, 40, 20, "")
  str_Left = StringGadget(#PB_Any, 230, 140, 40, 20, "")
  str_Right = StringGadget(#PB_Any, 230, 170, 40, 20, "")
  TopMargin.i = 15
  LeftMargin.i = 15
  BottomMargin.i = 15
 RightMargin.i = 15

  SetGadgetText(str_Top,Str(TopMargin.i))
  SetGadgetText(str_Left,Str(LeftMargin.i))
  SetGadgetText(str_Bottom,Str(BottomMargin.i))
  SetGadgetText(str_Right,Str(RightMargin.i))
  
  Repeat
    
    Event = WaitWindowEvent()

     If wnd_PageSetup > 0 
     Select event
    
       Case #PB_Event_Menu
         
         Select EventMenu()
             
         EndSelect
         
       Case #PB_Event_Gadget
         
         Select EventGadget()
                       
           Case btn_Ok
             Quit = #True
             CloseWindow(wnd_PageSetup)
     
           Case str_Top
             TopMargin.i = Val(GetGadgetText(str_Top))
            
           Case str_Left
             LeftMargin.i = Val(GetGadgetText(str_Left))          

             
           Case str_Bottom
            BottomMargin.i = Val(GetGadgetText(str_Bottom))          

             
           Case str_Right
             RightMargin.i = Val(GetGadgetText(str_Right))             

             EndSelect  

                       
         EndSelect
         
     EndSelect
   EndIf
   
 Until Quit = #True
 
   EndProcedure

EndModule
; IDE Options = PureBasic 5.51 (Windows - x64)
; CursorPosition = 14
; FirstLine = 67
; Folding = -
; EnableXP
; EnableUnicode