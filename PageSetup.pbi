;**********************************************************************
;
;    Page Setup Module For PB 5.4
;
;    FileName PageSetupfrm.pb
;
;******************************************************************
DeclareModule PageSetup
  
  Declare Open()
  Global TopMargin.i,LeftMargin.i,BottomMargin.i,RightMargin.i
  
EndDeclareModule

Module PageSetup

  Global btn_Ok,frm_Margins, str_Top, str_Left, str_Bottom, str_Right.i

  Procedure Open()
    
  wnd_PageSetup = OpenWindow(#PB_Any, x, y, 330, 160, "Set Margins For Printing", #PB_Window_TitleBar | #PB_Window_Tool | #PB_Window_ScreenCentered)

  FrameGadget(#PB_Any, 10, 10, 310, 100, " Margins ")
  TextGadget(#PB_Any, 130, 40, 30, 20, "mm")
  TextGadget(#PB_Any, 130, 70, 30, 20, "mm")
  TextGadget(#PB_Any, 170, 40, 50, 20, "Left", #PB_Text_Right)
  TextGadget(#PB_Any, 20, 40, 50, 20, "Top", #PB_Text_Right)
  TextGadget(#PB_Any, 280, 40, 30, 20, "mm")
  TextGadget(#PB_Any, 280, 70, 30, 20, "mm")
  TextGadget(#PB_Any, 20, 70, 50, 20, "Bottom", #PB_Text_Right)
  TextGadget(#PB_Any, 170, 70, 50, 20, "Right", #PB_Text_Right)
  str_Top = StringGadget(#PB_Any, 80, 40, 40, 20, "")
  str_Bottom = StringGadget(#PB_Any, 80, 70, 40, 20, "")
  str_Left = StringGadget(#PB_Any, 230, 40, 40, 20, "")
  str_Right = StringGadget(#PB_Any, 230, 70, 40, 20, "")
  btn_Ok = ButtonGadget(#PB_Any, 220, 120, 100, 30, "Ok")

  SetGadgetText(str_Top,Str(TopMargin.i))
  SetGadgetText(str_Left,Str(LeftMargin.i))
  SetGadgetText(str_Bottom,Str(BottomMargin.i))
  SetGadgetText(str_Right,Str(RightMargin.i))
  
  Repeat
    
    Event = WaitWindowEvent()

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

   
 Until Quit = #True
 
   EndProcedure

EndModule
; IDE Options = PureBasic 5.51 (Windows - x64)
; CursorPosition = 36
; FirstLine = 27
; Folding = -
; EnableXP
; EnableUnicode