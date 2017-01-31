DeclareModule SelectDate
  
  Global BirthDay.i,BirthMonth.i,BirthYear.i
  
  Declare.i Open()
  Declare.s NumberToMonth(SelMonth.i)
  
EndDeclareModule

Module SelectDate
  
  Global SelectedDate.i
  
  Enumeration 300
    #dlgDate
    #txtDay
    #txtMonth
    #txtYear
    #cmbDay
    #cmbMonth
    #cmbYear
    #btnOk
  EndEnumeration
  
  Procedure ShowFormTexts()
    
    SetWindowTitle(#dlgDate,"Select Date")
    SetGadgetText(#btnOk,"Ok")
    SetGadgetText(#txtDay,"Day")
    SetGadgetText(#txtMonth,"Month")
    SetGadgetText(#txtYear,"Year")

  EndProcedure
  
    Procedure.s NumberToMonth(SelMonth.i)
   
    Define RetVal.s
  
    Select SelMonth

      Case 1
        RetVal = "January"
      Case 2
        RetVal = "February"
      Case 3
        RetVal = "March"
      Case 4
        RetVal = "April"
      Case 5
        RetVal = "May"
      Case 6
        RetVal = "June"
      Case 7
        RetVal = "July"
      Case 8
        RetVal = "August"
      Case 9
        RetVal = "September"
      Case 10
        RetVal = "October"
      Case 11
        RetVal = "November"
      Case 12
        RetVal = "December"
        
    EndSelect
    
   ProcedureReturn RetVal
      
  EndProcedure
     
    Procedure LoadCombos()
      
      Define iLoop.i
      
      For iLoop = 1 To 31
        AddGadgetItem(#cmbDay,-1,Str(iLoop))
      Next iLoop
      
      For iLoop = 0 To 11
        
        AddGadgetItem(#cmbMonth,iLoop,NumberToMonth(iLoop + 1))
        SetGadgetItemData(#cmbMonth, iLoop, iLoop + 1)
      Next iLoop
      
      For iLoop = Year(Date()) - 100 To Year(Date())
        AddGadgetItem(#cmbYear,-1,Str(iLoop))
      Next iLoop
      
    EndProcedure
    
    Procedure.i Open()
    
    Define Date.i 
    Define Quit.i = #False
    OpenWindow(#dlgDate, 0, 0, 240, 120, "Select Date", #PB_Window_SystemMenu)
    TextGadget(#txtDay, 10, 10, 50, 20, "Day")
    TextGadget(#txtMonth, 90, 10, 60, 20, "Month")
    TextGadget(#txtYear, 170, 10, 60, 20, "Year")
    ComboBoxGadget(#cmbDay, 10, 40, 50, 20)
    ComboBoxGadget(#cmbMonth, 80, 40, 70, 20)
    ComboBoxGadget(#cmbYear, 170, 40, 60, 20)
    ButtonGadget(#btnOk, 160, 80, 70, 30, "Ok")
    
    ShowFormTexts()
    LoadCombos()
    
    SetGadgetText(#cmbDay,Str(BirthDay))
    SetGadgetText(#cmbMonth,NumberToMonth(BirthMonth))
    SetGadgetText(#cmbYear,Str(Birthyear))
    
    Repeat
      
      Event = WaitWindowEvent()
      
      Select  EventGadget()
          
          
        Case #btnOk
          
          BirthDay = Val(GetGadgetText(#cmbDay))
          BirthMonth = GetGadgetItemData(#cmbMonth,GetGadgetState(#cmbMonth))
          BirthYear = Val(GetGadgetText(#cmbYear))         
          CloseWindow(#dlgDate)
          Quit = #True  

      EndSelect      
          
    Until Quit = #True
    
  EndProcedure
  
  EndModule  
; IDE Options = PureBasic 5.51 (Windows - x64)
; CursorPosition = 117
; FirstLine = 43
; Folding = j-
; EnableXP
; EnableUnicode