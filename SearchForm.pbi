DeclareModule SearchFrm
  
  Declare.s Open()
  
EndDeclareModule

Module SearchFrm
  
  Enumeration 400
    #txtFirstName
    #strFirstName
    #txtLastName
    #strLastName
    #txtCountry
    #cmbCountry
  EndEnumeration
  
    Global WinSearch.i, txtFirstName, strFirstName, txtLastName.i,strLastName.i,btnOk, btnCancel

  Procedure LoadCombos()
      
    Define AddressDB.i,Criteria.s
    
    AddressDB = OpenDatabase(#PB_Any,GetCurrentDirectory() + "Database\Address Book.db","","")
    
    ;Country  
    Criteria = "SELECT DISTINCT Country From Contacts" 
    DatabaseQuery(AddressDB, Criteria)

    While NextDatabaseRow(AddressDB)
      
      AddGadgetItem(#cmbCountry,-1,GetDatabaseString(AddressDB,0))
      
    Wend
    
    FinishDatabaseQuery(AddressDB)
    CloseDatabase(AddressDB)   
      
  EndProcedure
    
  Procedure.s BuildSearchString()
   
    Define Criteria.s
    Define Built.i
  
    ;Build SQL WHERE clause
    Built = #False
    Criteria = ""
  
    ;First Name
    If Len(Trim(GetGadgetText(#strFirstName))) > 0
      Criteria = " WHERE FirstName Like '%" + GetGadgetText(#strFirstName) + "%'"
      Built = #True  
    EndIf
  
    ;Last Name
    If Len(Trim(GetGadgetText(#strLastName))) > 0
      If built = #True
        Criteria = Criteria + " AND LastName Like '%" + GetGadgetText(#strLastName) + "%'"
      Else
        Criteria = Criteria + " WHERE LastName Like '%" + GetGadgetText(#strLastName) + "%'"
        Built = #True 
      EndIf
    EndIf 
  
    ;Country
    If Len(Trim(GetGadgetText(#cmbCountry))) > 0
      If built = #True
        Criteria = Criteria + " AND Country = '" + GetGadgetText(#cmbCountry) + "'"
      Else
        Criteria = Criteria + " WHERE Country = '" + GetGadgetText(#cmbCountry) + "'"
        Built = #True 
      EndIf
    EndIf   
  
    ProcedureReturn Criteria
  
  EndProcedure 
 
  Procedure.s Open()

    Define Criteria.s
  
    WinSearch = OpenWindow(#PB_Any, 0, 0, 560, 80, "Enter Search Criteria", #PB_Window_SystemMenu | #PB_Window_WindowCentered)
    TextGadget(#txtFirstName, 10, 10, 100, 20, "First Name", #PB_Text_Right)
    StringGadget(#strFirstName, 120, 10, 150, 20, "")
    TextGadget(#txtLastName, 290, 10, 100, 20, "Last Name", #PB_Text_Right)
    StringGadget(#strLastName, 400, 10, 150, 20, "") 
    TextGadget(#txtCountry, 10, 40, 100, 20, "Country", #PB_Text_Right)  
    ComboBoxGadget(#cmbCountry, 120, 40, 110, 20)
    btnOk = ButtonGadget(#PB_Any, 310, 40, 60, 30, "Ok")
    btnCancel = ButtonGadget(#PB_Any, 410, 40, 60, 30, "Cancel")
  
    LoadCombos()
    
    Repeat
    
      Event = WaitWindowEvent()
  
      Select Event
          
        Case #PB_Event_CloseWindow
      
          Criteria = ""
          CloseWindow(WinSearch)
          Quit = #True

        Case #PB_Event_Gadget
      
          Select EventGadget()
          
            Case   btnOk
          
              Quit = #True
              Criteria = BuildSearchString()
              CloseWindow(WinSearch)
          
            Case btnCancel
          
              Criteria = ""
              CloseWindow(WinSearch)
              Quit = #True

          EndSelect
      
      EndSelect
  
    Until Quit = #True 

    ProcedureReturn Criteria
  
  EndProcedure

EndModule
; IDE Options = PureBasic 5.60 Beta 1 (Windows - x64)
; CursorPosition = 83
; FirstLine = 21
; Folding = y
; EnableXP
; EnableUnicode