;
; ------------------------------------------------------------
;
;   PureBasic - Print/Preview Module
;
;   FileName CDPrint.pbi
;
;   
; ------------------------------------------------------------
;
;UseJPEGImageDecoder()
DeclareModule CDPrint
  
  ;Page Variables This will change as I proceed to maybe an array of a page structure
  Global  PageHeight.i ;Height of Selected Page in mm
  Global  PageWidth.i  ;Width of Selected Page in mm
  Global  TopMargin.i  ;Top Margin In mm
  Global  LeftMargin.i ;Left Margin In mm
  Global  BottomMargin.i ;Bottom Margin In mm
  Global  RightMargin.i  ;Right Margin In mm
  
 ;Experimenting with Orientation to attempt to allow a mix of portrait and Landscape in the same document 
  Global  Orientation.i = -1 ;Portrait or Landscape
  
  Global  Preview.i
  GPageHeight.i ;mm To work out Graphic scale factor
  GPageWidth.i  ;mm To work out Graphic scale factor
  Orientation.i ;Portrait or Landscape for print routines
  TPageHeight.i ;Points To work out Text scale factor
  TPageWidth.i  ;Points To work out Text scale factor
  
  Declare CDPrintEvents(Event)
  Declare.i open(JobName.s)
  Declare ShowPage(PageID)
  Declare AddPage()
  Declare Finished()
  Declare PrintLine(Startx,Starty,Endx,Endy,LineWidth)
  Declare PrintBox(Topx,Topy,Bottomx,Bottomy,LineWidth)
  Declare PrintText(Startx,Starty,Font.s,Size.i,Text.s)
  Declare PrintImage(Image.l,Topx.i,Topy.i,Width.i,Height.i)
  Declare.f GettextWidthmm(text.s,FName.s,FSize.f)
  Declare.f GettextHeightmm(text.s,FName.s,FSize.f)
  ;Declare Page_Setup()
  
EndDeclareModule

Module CDPrint
  
  ;Structure For Usefull Printer Information
  Structure PInfo
    TopMargin.i
    LeftMargin.i
    BottomMargin.i
    RightMargin.i
    Width.i
    Height.i
    UsableWidth.i
    UsableHeight.i
  EndStructure

  Global PrinterInfo.PInfo
  Global PreviewWindow.l
  Global Dim PageNos.i(1)
  Global PageNo.i
  Global GraphicScale.f
  Global TextScale.f  
  Global Result
  
  ;PageSetup Form Variables
  Global frm_Printer, btn_Cancel, btn_Ok, cmb_Printers, cmb_PaperSize, frm_Margins, str_Top, str_Left, String_0, str_Bottom, str_Right, String_0_Copy3, img_Back, img_Page, opt_Portrait, opt_Landscape, frm_Orientation
  
  Procedure GetPrinterInfo()
  
  CompilerSelect #PB_Compiler_OS
    
    CompilerCase   #PB_OS_MacOS
      
      ;The vectordrawing functions print correctly on the MAC so simply set all to zero
      PrinterInfo\Width = 0
      PrinterInfo\Height = 0
      PrinterInfo\UsableWidth = 0
      PrinterInfo\UsableHeight = 0
      PrinterInfo\TopMargin = 0
      PrinterInfo\LeftMargin = 0
      PrinterInfo\BottomMargin = 0
      PrinterInfo\RightMargin = 0

    CompilerCase   #PB_OS_Linux   
      
      ;Not Defined Yet
      
    CompilerCase   #PB_OS_Windows   
      
      Define HDPmm.d
      Define VDPmm.d
        
      printer_DC.l = StartDrawing(PrinterOutput())

      If printer_DC
        HDPmm = GetDeviceCaps_(printer_DC,#LOGPIXELSX) / 25.4
        VDPmm = GetDeviceCaps_(printer_DC,#LOGPIXELSY) / 25.4
        PrinterInfo\Width = GetDeviceCaps_(printer_DC,#PHYSICALWIDTH) / HDPmm
        PrinterInfo\Height = GetDeviceCaps_(printer_DC,#PHYSICALHEIGHT) / VDPmm
        PrinterInfo\UsableWidth = GetDeviceCaps_(printer_DC,#HORZSIZE) ; Returns mm
        PrinterInfo\UsableHeight = GetDeviceCaps_(printer_DC,#VERTSIZE); Returns mm
        PrinterInfo\TopMargin = GetDeviceCaps_(printer_DC,#PHYSICALOFFSETY) / VDPmm
        PrinterInfo\LeftMargin = GetDeviceCaps_(printer_DC,#PHYSICALOFFSETX) / HDPmm
        PrinterInfo\BottomMargin = PrinterInfo\Height - PrinterInfo\UsableHeight - PrinterInfo\TopMargin
        PrinterInfo\RightMargin = PrinterInfo\Width - PrinterInfo\UsableWidth - PrinterInfo\LeftMargin
      EndIf

      StopDrawing()
      
  CompilerEndSelect
   
EndProcedure
  
  Procedure.f GettextWidthmm(text.s,FName.s,FSize.f)
    
    LoadFont(0,FName.s, FSize.f)    ;Load Font In Points
    ASize.f = FSize.f * 0.352777778 ;Convert Font Points To mm
    VectorFont(FontID(0), ASize.f ) ;Use Font In mm Size
    ProcedureReturn VectorTextWidth(text.s,#PB_VectorText_Visible) ;Width of text In mm

  EndProcedure
  
  Procedure.f GettextHeightmm(text.s,FName.s,FSize.f)
    
    LoadFont(0,FName.s, FSize.f)    ;Load Font In Points
    ASize.f = FSize.f * 0.352777778 ;Convert Font Points To mm
    VectorFont(FontID(0), ASize.f)  ;Use Font In mm Size
    ProcedureReturn VectorTextHeight(text.s,#PB_VectorText_Visible) ;Height of text In mm

  EndProcedure 
  
  Procedure ShowPage(PageID)
    
    If CDPrint::Preview.i = 1
      SetGadgetState(66,PageID + 1)
      SetGadgetState(34, ImageID(PageNos.i(PageID)))
      Top = 25
      Left = 10
      ResizeGadget(34,Left,Top,#PB_Ignore,#PB_Ignore)
    EndIf  
    ResizeGadget(66,GadgetWidth(34) - 30,0,#PB_Ignore,#PB_Ignore)
    ResizeWindow(PreviewWindow,#PB_Ignore,#PB_Ignore,GadgetWidth(34) + 20,GadgetHeight(34) + 35)
    
  EndProcedure
  
  Procedure.i Open(JobName.s)
    
    If PrintRequester()
      
      GetPrinterInfo()
      PageHeight = PrinterInfo\Height
      PageWidth = PrinterInfo\Width

      TPageHeight.i = PageHeight * 2.834645669 ;mm To Points
      TPageWidth.i = PageWidth * 2.834645669

      If PageHeight > GPageWidth.i
        Orientation = 1
        GraphicScale.f = 500/PageHeight
        TextScale.f = 500/TPageHeight.i
      Else
        Orientation = 0
        GraphicScale.f = 500/PageWidth
        TextScale.f = 500/TPagewidth.i
      EndIf
  
      If Preview.i = 1 ;If preview mode open Preview Window
    
        PreviewWindow.l = OpenWindow(#PB_Any, #PB_Ignore,#PB_Ignore, 500, 535, "Print Preview - " + JobName.s)
      
        If PreviewWindow.l > 0
        
          SpinGadget     (66, 450, 0, 50, 25, 0, 1000,#PB_Spin_Numeric)
          SetGadgetState (66, 1)
          ImageGadget(34, 5, 5, 50, 50,  0,#PB_Image_Raised)

          ;Set Page Counter To Zero And Create first Page Image
          PageNo.i = 0 
        
          PageNos.i(PageNo.i) = CreateImage(#PB_Any, PageWidth * GraphicScale.f,PageHeight * GraphicScale.f, 32,RGB(255,255,255))
          ShowPage(PageNo.i)
        
          StartVectorDrawing(ImageVectorOutput(PageNos.i(PageNo.i) ))
        
          ;Set all printer margins to zero for preview
          PrinterInfo\TopMargin = 0
          PrinterInfo\LeftMargin = 0
          PrinterInfo\BottomMargin = 0
          PrinterInfo\RightMargin = 0
          ProcedureReturn PreviewWindow.l
        
        Else
      
          ;Tell The User We Failed
      
          ProcedureReturn 0
 
        EndIf
    
      Else
    
        ;Start Printing
        Result = StartPrinting(JobName.s)
        StartVectorDrawing(PrinterVectorOutput(#PB_Unit_Millimeter)) 

        If Orientation = 0
          RotateCoordinates(0,0,-90)
          TranslateCoordinates(-PrinterInfo\Height, 0)
        EndIf

        ProcedureReturn result

      EndIf
    Else
      
      ;User Cancelled Printer selection
      ProcedureReturn 0
      
    EndIf
  
  EndProcedure
  
  Procedure AddPage()
    
    If Preview.i = 1
      
      ;Increase Last Page Number  
      PageNo = PageNo + 1
      
      ;Increase Handle Array by one
      ReDim PageNos(PageNo)

      ;Show new page number in spin gadget
      
      SetGadgetState(66,PageNo.i + 1)
      
      SetGadgetText(66, Str(PageNo.i + 1))

      ;Create the image for the page and store handle in array
      PageNos.i(PageNo.i) = CreateImage(#PB_Any, PrinterInfo\Width * GraphicScale.f,PrinterInfo\Height * GraphicScale.f, 32,RGB(255,255,255))

      StartVectorDrawing(ImageVectorOutput(PageNos.i(PageNo.i) ))
       
    Else
      
      ;Printer New Page
      NewPrinterPage()
      If Orientation = 0
        RotateCoordinates(0,0,-90)
        TranslateCoordinates(-297, 0)
      EndIf
    EndIf
    
  EndProcedure
  
  Procedure PrintLine(Startx,Starty,Endx,Endy,LineWidth)
    
    If Preview.i = 1
      
      ;Scaled To Fit Preview Window   
      MovePathCursor(Startx * GraphicScale.f, Starty * GraphicScale.f)
        
      AddPathLine(Endx * GraphicScale.f, Endy * GraphicScale.f, #PB_Path_Default)
        
      VectorSourceColor(RGBA(0, 0, 0, 255))
        
      StrokePath(LineWidth * GraphicScale.f, #PB_Path_RoundCorner)

    Else
      
      ;Print Routine No Scaling but printer hard margins
      Startx = Startx - PrinterInfo\LeftMargin
      Starty = Starty - PrinterInfo\TopMargin
      Endx = Endx - PrinterInfo\LeftMargin
      Endy = Endy - PrinterInfo\TopMargin  
      MovePathCursor(Startx, Starty)
        
      AddPathLine(Endx, Endy, #PB_Path_Default)
        
      VectorSourceColor(RGBA(0, 0, 0, 255))
        
      StrokePath(LineWidth,#PB_Path_RoundCorner)

    EndIf

  EndProcedure
 
  Procedure PrintBox(Topx,Topy,Bottomx,Bottomy,LineWidth)
    
    If Preview.i = 1
      
      ;Scaled To Fit Preview Image

      AddPathBox(Topx * GraphicScale.f, Topy * GraphicScale.f, (Bottomx - Topx) * GraphicScale.f, (Bottomy - Topy) * GraphicScale.f)

      VectorSourceColor(RGBA(255, 0, 0, 255))
    
      StrokePath(LineWidth * GraphicScale.f)
 
    Else
      
      ;Print Routine. No Scaling
      Topx = Topx - PrinterInfo\LeftMargin
      Topy = Topy - PrinterInfo\TopMargin
      Bottomx = Bottomx - PrinterInfo\LeftMargin
      Bottomy = Bottomy - PrinterInfo\TopMargin     
      AddPathBox(Topx, Topy , (Bottomx - Topx), (Bottomy - Topy))

      VectorSourceColor(RGBA(255, 0, 0, 255))
            
      StrokePath(LineWidth)
  
  EndIf 
    
  EndProcedure
  
  Procedure PrintText(Startx,Starty,Font.s,Size.i,Text.s)
    
    ;Scaled To Fit Preview Image
    
    If Preview.i = 1
      
      LoadFont(0, Font.s , Size.i )
      
      ASize.f = Size.i * 0.352777778 ;Convert Font Points To mm
      
      MovePathCursor(Startx * GraphicScale.f, Starty * GraphicScale.f)

      VectorFont(FontID(0), ASize.f * GraphicScale.f)
 
      VectorSourceColor(RGBA(0, 0, 0, 255))

      DrawVectorText(Text.s)
      
    Else
      
      ;Print Routine. No Scaling
      Startx = Startx - PrinterInfo\LeftMargin
      Starty = Starty - PrinterInfo\TopMargin
      LoadFont(0, Font.s , Size.i )
      ASize.f = Size.i * 0.352777778 ;Convert Font Points To mm
      VectorFont(FontID(0), ASize.f)
      VectorSourceColor(RGBA(0, 0, 0, 255))

      MovePathCursor(Startx, Starty)
      DrawVectorText(Text.s)
      
    EndIf

  EndProcedure

  Procedure PrintImage(Image.l,Topx.i,Topy.i,Width.i,Height.i)
    
    If Preview.i = 1

        MovePathCursor(Topx * GraphicScale.f, Topy * GraphicScale.f)
 
        DrawVectorImage(ImageID(Image.l),100,Width.i * GraphicScale.f,Height.i * GraphicScale.f)
 
    Else
            
      ;Print Routine
      Topx = Topx - PrinterInfo\LeftMargin
      Topy = Topy - PrinterInfo\TopMargin
      MovePathCursor(Topx, Topy)
      
      DrawVectorImage(ImageID(Image),255,Width,Height)
 
  EndIf
  
  EndProcedure

  Procedure Finished()
    
    If Preview.i = 0
      StopVectorDrawing()
      StopPrinting()
    Else
      StopVectorDrawing()
    EndIf
        
  EndProcedure
  
  Procedure CDPrintEvents(Event)
    
    If event = #PB_Event_CloseWindow
      CloseWindow(PreviewWindow.l)
    EndIf
      
    If Event = #PB_Event_Gadget

      Select EventGadget()
        Case 66
          
          SetGadgetText(66, Str(GetGadgetState(66)))
          If GetGadgetState(66) > 0 And GetGadgetState(66) -1 <= PageNo.i
            ShowPage(GetGadgetState(66) -1)
          ElseIf GetGadgetState(66) < 1
            ;Show Last Page No or first page Number in gadget
            SetGadgetState(66,1)
          Else
            SetGadgetState(66,PageNo.i + 1 )
          EndIf
          
        EndSelect       
         
        EndIf
    
  EndProcedure

EndModule
; IDE Options = PureBasic 5.60 Beta 1 (Windows - x64)
; CursorPosition = 213
; FirstLine = 81
; Folding = DD9
; EnableXP
; EnableUnicode