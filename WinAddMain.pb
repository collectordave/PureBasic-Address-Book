UseSQLiteDatabase()
UsePNGImageDecoder()
UseJPEGImageDecoder()
UseJPEGImageEncoder()

IncludeFile "DCTool.pbi"
IncludeFile "CDPrint.pbi"
IncludeFile "SearchForm.pbi"
IncludeFile "PageSetup.pbi"
IncludeFile "dlgSelectDate.pbi"
IncludeFile "dlgHelpViewer.pbi"

Enumeration MainWindow 100
  #mnuMain
  #mnuAdd
  #mnuEdit
  #mnuDelete
  #MenuItem_5
  #mnuSearch
  #mnuPageSetup
  #mnuPrint
  #mnuPrintBook
  #mnuPreview
  #mnuPreviewBook
  #MenuItem_13
  #mnuExit
  #cntimg
  #frmImage
  #imgContact
  #btnAssign
  #btnDelete
  #frmDetail
  #txtFristName
  #strFirstName
  #txtLastName
  #strLastName
  #txtAddress
  #strAddressLine1
  #strAddressLine2
  #strAddressLine3
  #txtPostCode
  #strPostCode
  #txtTelephone
  #strTelephone
  #txtMobile
  #strMobile
  #txtEMail
  #strEMail
  #txtCountry
  #strCountry
  #txtBirthday
  #strBirthday
  #btnBirthday
  #chkFamily
  #chkFriend
  #chkOther
  #imgFirst
  #imgPrevious
  #imgNext
  #imgLast
  #btnFirst
  #btnPrevious
  #btnNext 
  #btnLast
  #imgAdd
  #imgEdit
  #imgDelete
  #imgExit
  #imgOk
  #imgCancel
  #imgHelp
  #imgPrint
  #imgSearch
  #Print
EndEnumeration

Global AddressMain.i,ToolBarMain.i,ToolBarEdit.i,AddressDB.i,TotalRows.i,CurrentContactID.i,CurrentRow.i
Global SearchCriteria.s,CBirthDay.i,CBirthMonth.i,CBirthYear.i
Global PrintControl.i,Mode.s

CatchImage(#imgFirst,?FirstPhoto)
CatchImage(#imgPrevious,?PreviousPhoto)
CatchImage(#imgNext,?NextPhoto)
CatchImage(#imgLast,?LastPhoto)
CatchImage(#imgAdd,?ToolBarAdd)
CatchImage(#imgEdit,?ToolBarEdit)
CatchImage(#imgDelete,?ToolBarDelete)
CatchImage(#imgExit,?ToolBarExit)
CatchImage(#imgHelp,?ToolBarHelp)
CatchImage(#imgPrint,?ToolBarPrint)
CatchImage(#imgSearch,?ToolBarSearch)
CatchImage(#imgOk,?ToolBarOk)
CatchImage(#imgCancel,?ToolBarCancel)

;Default Page Margins
PageSetup::TopMargin.i = 10
PageSetup::LeftMargin.i = 10
PageSetup::BottomMargin.i = 10
PageSetup::RightMargin.i = 10

Procedure OpenAddressDB()
  
  AddressDB = OpenDatabase(#PB_Any,GetCurrentDirectory() + "Database\Address Book.db","","")
  
EndProcedure

Procedure ClearDisplay()
  
  ;Clear All Display Boxes
  SetGadgetText(#strFirstName,"")
  SetGadgetText(#strLastName,"")
  SetGadgetText(#strAddressLine1,"")
  SetGadgetText(#strAddressLine2,"")
  SetGadgetText(#strAddressLine3,"")
  SetGadgetText(#strPostCode,"")
  SetGadgetText(#strCountry,"")
  SetGadgetText(#strBirthday,"") 
  SetGadgetText(#strEMail,"")
  SetGadgetText(#strTelephone,"")
  SetGadgetText(#strMobile,"")
  SetGadgetState(#imgContact,0)  
  CBirthday = 1
  CBirthMonth = 1
  CBirthYear = 2000
  
EndProcedure

Procedure GetTotalRecords()
  
  Define iLoop.i
  Define GetTotal.s,SearchString.s
  
  ;Find out how many records will be returned
  TotalRows = 0
  SearchString = "SELECT * FROM Contacts " + SearchCriteria
  
  If DatabaseQuery(AddressDB, SearchString)
    
    While NextDatabaseRow(AddressDB)

      TotalRows = TotalRows + 1
      
    Wend
    
    FinishDatabaseQuery(AddressDB)  

  EndIf

EndProcedure

Procedure CheckRecords()

  ;Sort out the navigation buttons
  If TotalRows < 2
    
    ;Only one record so it is the first and the last
    DisableGadget(#btnLast, #True)     ;No move last as allready there
    DisableGadget(#btnNext, #True)     ;No next record as this is the last record
    DisableGadget(#btnFirst, #True)    ;No first record as this is the first record
    DisableGadget(#btnPrevious, #True) ;No previous record as this is the first record
    
  ElseIf CurrentRow = 1
    ;On the first row with more than one selected
    DisableGadget(#btnLast, 0)     ;Can move to last record
    DisableGadget(#btnNext, 0)     ;Can move to next record
    DisableGadget(#btnFirst, #True)    ;No first record as this is the first record
    DisableGadget(#btnPrevious, #True) ;No previous record as this is the first record
    
  ElseIf  CurrentRow = TotalRows
    
    ;If on the last record
    DisableGadget(#btnLast, #True)     ;No move last as allready there
    DisableGadget(#btnNext, #True)     ;No next record as this is the last record
    DisableGadget(#btnFirst, 0)    ;Can still move to first record
    DisableGadget(#btnPrevious, 0) ;Can still move to previous record
    
  Else
    
    ;Somewhere in the middle of the selected records
    DisableGadget(#btnLast, 0)     ;Can move to last record
    DisableGadget(#btnNext, 0)     ;Can move to next record
    DisableGadget(#btnFirst, 0)    ;Can move to first record
    DisableGadget(#btnPrevious, 0) ;Can move to previous record
    
  EndIf

EndProcedure

Procedure GetImageFromDB()
  
  Define PictureSize.i,Picture.i,x.i,y.i
  
  PictureSize = DatabaseColumnSize(AddressDB, DatabaseColumnIndex(AddressDB,"Thumb"))
  If PictureSize > 0
    Picture = AllocateMemory(PictureSize)
    GetDatabaseBlob(AddressDB, DatabaseColumnIndex(AddressDB,"Thumb"), Picture, PictureSize)
    CatchImage(54, Picture, PictureSize)
    FreeMemory(Picture)
    x = (200 - ImageWidth(54))/2
    y = (200 - ImageHeight(54))/2
    ResizeGadget(#imgContact,x,y,ImageWidth(54),ImageHeight(54))
    SetGadgetState(#imgContact,ImageID(54))
  Else
    SetGadgetState(#imgContact,0)    
  EndIf
  
EndProcedure

Procedure AddContact()
  
  Define Criteria.s

  Criteria = "INSERT INTO Contacts (FirstName) VALUES ('" + ReplaceString(GetGadgetText(#strFirstName),"'","''") + "')"
  DatabaseUpdate(AddressDB,Criteria)
  
  Criteria = "Select last_insert_rowid()"
  
  DatabaseQuery(AddressDB,Criteria)
  
  FirstDatabaseRow(AddressDB)
  
  CurrentContactID =  GetDatabaseLong(AddressDB,0)
  
  FinishDatabaseQuery(AddressDB)
  
EndProcedure

Procedure DeleteContact()
  
  Define Criteria.s
  
  Criteria = "DELETE FROM Contacts WHERE Contact_ID = " + Str(CurrentContactID) + ";"
  DatabaseUpdate(AddressDB,Criteria)
  
  Criteria = "Vacuum"
  DatabaseUpdate(AddressDB,Criteria)
  
EndProcedure

Procedure SaveContact()
  
  Define SaveCriteria.s
  
  SaveCriteria = "UPDATE Contacts SET " + 
                 "FirstName = '" + ReplaceString(GetGadgetText(#strFirstName),"'","''") + 
                 "',LastName ='" + ReplaceString(GetGadgetText(#strLastName),"'","''") +  
                 "',Addr1 ='" + ReplaceString(GetGadgetText(#strAddressLine1),"'","''") + 
                 "',Addr2 ='" + ReplaceString(GetGadgetText(#strAddressLine2),"'","''") +                  
                 "',Addr3 ='" + ReplaceString(GetGadgetText(#strAddressLine3),"'","''") +   
                 "',Country ='" + ReplaceString(GetGadgetText(#strCountry),"'","''") + 
                 "',PostCode ='" + ReplaceString(GetGadgetText(#strPostCode),"'","''") + 
                 "',Telephone='" + ReplaceString(GetGadgetText(#strTelephone),"'","''") + 
                 "',Mobile ='" + ReplaceString(GetGadgetText(#strMobile),"'","''") + 
                 "',EMail='" + ReplaceString(GetGadgetText(#strEMail),"'","''") +                  
                 "',BirthDay = " + Str(CBirthDay) +   
                 " ,BirthMonth = " + Str(CBirthMonth) +  
                 " ,BirthYear = " + Str(CBirthYear) +  
                 " WHERE Contact_ID = " + Str(CurrentContactID) + ";"

  DatabaseUpdate(AddressDB,SaveCriteria)

  
  
EndProcedure

Procedure AddThumb()
    
  Structure ImageData
    *address
    size.i
  EndStructure
    
  Define ImageToSave.i,RetVal.i
  Define AspectRatioWidth.f,AspectRatioHeight.f
  Define ImageFolder.s,PhotoFileName.s,Criteria.s
    
  PhotoFileName = OpenFileRequester("Select Image", "C:\PB Projects\Address Book\Images","Image (*.jpg)|*.jpg", 0) 
  If PhotoFileName
    
    
    ;Resize Image To Thumbnail size
    ImageToSave = LoadImage(#PB_Any, PhotoFileName)

    AspectRatioWidth = 200/ImageWidth(ImageToSave)
    AspectRatioHeight = 200/ImageHeight(ImageToSave)   
    If AspectRatioWidth < AspectRatioHeight
      ResizeImage(ImageToSave,ImageWidth(ImageToSave) * AspectRatioWidth,ImageHeight(ImageToSave) * AspectRatioWidth)
    Else
      ResizeImage(ImageToSave,ImageWidth(ImageToSave) * AspectRatioHeight,ImageHeight(ImageToSave) * AspectRatioHeight)
    EndIf
    newImg = EncodeImage(ImageToSave,#PB_ImagePlugin_JPEG)
    NewImgSize = MemorySize(newImg)

    SetDatabaseBlob(AddressDB, 0, NewImg, NewImgSize)
    Criteria = "UPDATE Contacts SET Thumb = ?  WHERE Contact_ID = " + Str(CurrentContactID)
   
    DatabaseUpdate(AddressDB, Criteria)     
  EndIf
  
  EndProcedure

Procedure DisplayRecord()
      
  Define SearchString.s,SMonth.s
  Define DBYear.i,DBMonth.i

  SearchString = "SELECT * FROM Contacts " + SearchCriteria + " ORDER BY LastName ASC LIMIT 1 OFFSET " + Str(CurrentRow -1)

  DatabaseQuery(AddressDB, SearchString)

  If FirstDatabaseRow(AddressDB)
    CurrentContactID = GetDatabaseLong(AddressDB, DatabaseColumnIndex(AddressDB, "Contact_ID"))
    SetGadgetText(#strFirstName,GetDatabaseString(AddressDB, DatabaseColumnIndex(AddressDB, "FirstName")))
    SetGadgetText(#strLastName,GetDatabaseString(AddressDB, DatabaseColumnIndex(AddressDB, "LastName")))    
    SetGadgetText(#strAddressLine1,GetDatabaseString(AddressDB, DatabaseColumnIndex(AddressDB, "Addr1")))
    SetGadgetText(#strAddressLine2,GetDatabaseString(AddressDB, DatabaseColumnIndex(AddressDB, "Addr2")))    
    SetGadgetText(#strAddressLine3,GetDatabaseString(AddressDB, DatabaseColumnIndex(AddressDB, "Addr3")))   
    SetGadgetText(#strCountry,GetDatabaseString(AddressDB, DatabaseColumnIndex(AddressDB, "Country")))
    SetGadgetText(#strPostCode,GetDatabaseString(AddressDB, DatabaseColumnIndex(AddressDB, "PostCode")))    
    SetGadgetText(#strTelephone,GetDatabaseString(AddressDB, DatabaseColumnIndex(AddressDB, "Telephone")))       
    SetGadgetText(#strMobile,GetDatabaseString(AddressDB, DatabaseColumnIndex(AddressDB, "Mobile")))    
    SetGadgetText(#strEMail,GetDatabaseString(AddressDB, DatabaseColumnIndex(AddressDB, "EMail")))     
    ;Birthday
    CBirthDay = GetDatabaseLong(AddressDB, DatabaseColumnIndex(AddressDB, "BirthDay"))
    CBirthMonth = GetDatabaseLong(AddressDB, DatabaseColumnIndex(AddressDB, "BirthMonth"))
    CBirthYear = GetDatabaseLong(AddressDB, DatabaseColumnIndex(AddressDB, "BirthYear"))
    If CBirthDay > 0
      SetGadgetText(#strBirthday,Str(CBirthDay) + " " + SelectDate::NumberToMonth(CBirthMonth) + " " + Str(CBirthYear))
    Else
      SetGadgetText(#strBirthday,"")
    EndIf

    ;Family,friend etc
    
    ;Show picture if there is one
    GetImageFromDB()
    FinishDatabaseQuery(AddressDB)
    
  EndIf 
  
EndProcedure

Procedure.i PrintAddresses()
  
 
  Define Startx.i,Starty.i,PageHeight.i,PageWidth.i,MinimumX.i
  Define PictureSize.i,Picture.i
  Define SearchString.s

  PrintControl = CDPrint::open("Address Book")

  If PrintControl = 0
    
    ProcedureReturn 0
    
  EndIf
  
    PageHeight = CDPrint::PageHeight - PageSetup::TopMargin - PageSetup::BottomMargin
    PageWidth = CDPrint::PageWidth - Pagesetup::LeftMargin - Pagesetup::RightMargin
    
    ;Get contacts selected by user 
    SearchString = "SELECT * FROM Contacts " + SearchCriteria + "ORDER BY LastName ASC" ;LIMIT 1 OFFSET " + Str(CurrentRow -1)

    DatabaseQuery(AddressDB, SearchString)
  
    Startx = Pagesetup::LeftMargin
    Starty = PageSetup::TopMargin
    
    ;Now Cycle Through all selected contacts
    While NextDatabaseRow(AddressDB)
    
      If (Starty + 35) > PageHeight
      
        CDPrint::AddPage()
        Starty = PageSetup::TopMargin
    
      EndIf
    
      PictureSize = DatabaseColumnSize(AddressDB, DatabaseColumnIndex(AddressDB,"Thumb"))
      If PictureSize > 0

        Picture = AllocateMemory(PictureSize)
        GetDatabaseBlob(AddressDB, DatabaseColumnIndex(AddressDB,"Thumb"), Picture, PictureSize)
        CatchImage(54, Picture, PictureSize)
        FreeMemory(Picture)
        If ImageHeight(54) > ImageWidth(54)
          adjustedheight = 30 * (ImageHeight(54)/ImageWidth(54))
          adjustedwidth = adjustedheight * ImageWidth(54) / ImageHeight(54)
        Else
          adjustedwidth = 30 * (ImageWidth(54)/ImageHeight(54))
          adjustedheight = adjustedwidth * ImageHeight(54)/ImageWidth(54)
        EndIf
        CDPrint::PrintImage(54,Startx,Starty,adjustedwidth,adjustedheight) 
      EndIf
      Startx = Pagesetup::LeftMargin + 35
      CDPrint::PrintText(Startx,Starty,"Arial",12,GetDatabaseString(AddressDB, DatabaseColumnIndex(AddressDB, "FirstName")))
      MinimumX = 2 + Startx + CDPrint::GettextWidthmm(GetDatabaseString(AddressDB, DatabaseColumnIndex(AddressDB, "FirstName")),"Arial",12)
      CDPrint::PrintText(MinimumX,Starty,"Arial",12,GetDatabaseString(AddressDB, DatabaseColumnIndex(AddressDB, "LastName")))
      CDPrint::PrintText(Startx,Starty + 6,"Arial",12,GetDatabaseString(AddressDB, DatabaseColumnIndex(AddressDB, "Addr1")))
      CDPrint::PrintText(Startx,Starty + 12,"Arial",12,GetDatabaseString(AddressDB, DatabaseColumnIndex(AddressDB, "Addr2")))
      CDPrint::PrintText(Startx,Starty + 18,"Arial",12,GetDatabaseString(AddressDB, DatabaseColumnIndex(AddressDB, "Addr3")))
      CDPrint::PrintText(Startx,Starty + 24,"Arial",12,GetDatabaseString(AddressDB, DatabaseColumnIndex(AddressDB, "PostCode")))

      Startx = 30 + ((PageWidth - 40)/2)
      CDPrint::PrintText(Startx,Starty,"Arial",12,"Telephone")
      CDPrint::PrintText(startx + 40,Starty,"Arial",12,GetDatabaseString(AddressDB, DatabaseColumnIndex(AddressDB, "Telephone")))
      CDPrint::PrintText(startx,Starty + 6,"Arial",12,"Mobile")   
      CDPrint::PrintText(startx + 40,Starty + 6,"Arial",12,GetDatabaseString(AddressDB, DatabaseColumnIndex(AddressDB, "Mobile")))
      CDPrint::PrintText(startx,Starty + 12,"Arial",12,"EMail")    
      CDPrint::PrintText(startx + 40,Starty + 12,"Arial",12,GetDatabaseString(AddressDB, DatabaseColumnIndex(AddressDB, "EMail")))
 
      Startx = Pagesetup::LeftMargin
      Starty = Starty + 35

    Wend
  
    ;Finish The Print Job
    CDPrint::Finished()
    ProcedureReturn 1
    
EndProcedure

AddressMain = OpenWindow(#PB_Any, 0, 0, 740, 362, "Address Book", #PB_Window_SystemMenu|#PB_Window_ScreenCentered)

;Menu
CreateMenu(#mnuMain, WindowID(AddressMain))
  MenuTitle("Contacts")
  MenuItem(#mnuAdd, "Add")
  MenuItem(#mnuEdit, "Edit")
  MenuItem(#mnuDelete, "Delete")
  MenuBar()
  MenuItem(#mnuSearch, "Search")
  MenuBar()  
  MenuItem(#mnuPageSetup, "Page Setup")
  OpenSubMenu("Print")
  MenuItem(#mnuPrintBook, "Book")
  CloseSubMenu()
  OpenSubMenu("Print Preview")
  MenuItem(#mnuPreviewBook, "Book")
  CloseSubMenu()
  MenuBar()
  MenuItem(#mnuExit, "Exit")
  
;ToolBar
ToolBarMain = IconBarGadget(0, 0, WindowWidth(AddressMain),20,#IconBar_Default,AddressMain) 
AddIconBarGadgetItem(ToolBarMain, "Add", #imgAdd)
AddIconBarGadgetItem(ToolBarMain, "Edit", #imgEdit)
AddIconBarGadgetItem(ToolBarMain, "Delete", #imgDelete)
IconBarGadgetDivider(ToolBarMain)
AddIconBarGadgetItem(ToolBarMain, "Search", #imgSearch)
AddIconBarGadgetItem(ToolBarMain, "Print Preview", #imgPrint)
IconBarGadgetDivider(ToolBarMain)
AddIconBarGadgetItem(ToolBarMain, "Exit", #imgExit)
IconBarGadgetSpacer(ToolBarMain)
AddIconBarGadgetItem(ToolBarMain, "Help", #imgHelp)
ResizeIconBarGadget(ToolBarMain, #PB_Ignore, #IconBar_Auto)  
SetIconBarGadgetColor(ToolBarMain, 1, RGB(176,224,230))

ToolBarEdit = IconBarGadget(0, 0, WindowWidth(AddressMain),20,#IconBar_Default,AddressMain) 
AddIconBarGadgetItem(ToolBarEdit, "Ok", #imgOk)
AddIconBarGadgetItem(ToolBarEdit, "Cancel", #imgCancel)
ResizeIconBarGadget(ToolBarEdit, #PB_Ignore, #IconBar_Auto)  
SetIconBarGadgetColor(ToolBarEdit, 1, RGB(176,224,230))
HideIconBarGadget(ToolBarEdit,#True)

;Images
FrameGadget(#frmImage, 10, 45, 220, 280, " Image ")
ContainerGadget(#cntimg, 15, 60, 200, 200)
  SetGadgetColor(#cntimg, #PB_Gadget_BackColor,RGB(0,0,0))
  ImageGadget(#imgContact, 0, 0, 200, 200, 0)
CloseGadgetList()

ButtonGadget(#btnAssign, 20, 280, 60, 25, "Assign")
GadgetToolTip(#btnAssign, "Assign Image For Contact")
ButtonGadget(#btnDelete, 160, 280, 60, 25, "Delete")
GadgetToolTip(#btnDelete, "Delete Image")
  
;Contact Detail
FrameGadget(#frmDetail, 240, 45, 490, 280, " Contact Detail ")
TextGadget(#txtFristName, 250, 65, 100, 20, "First Name", #PB_Text_Right)
StringGadget(#strFirstName, 360, 65, 110, 20, "")
TextGadget(#txtLastName, 480, 65, 90, 20, "Last Name", #PB_Text_Right)
StringGadget(#strLastName, 580, 65, 130, 20, "")
TextGadget(#txtAddress, 260, 95, 100, 20, "Address")
StringGadget(#strAddressLine1, 260, 125, 210, 20, "")
StringGadget(#strAddressLine2, 260, 155, 210, 20, "")
StringGadget(#strAddressLine3, 260, 185, 210, 20, "")
TextGadget(#txtPostCode, 260, 210, 80, 20, "Post Code")
StringGadget(#strPostCode, 260, 235, 210, 20, "")
TextGadget(#txtTelephone, 490, 95, 100, 20, "Telephone")
StringGadget(#strTelephone, 490, 125, 220, 20, "")
TextGadget(#txtMobile, 490, 155, 100, 20, "Mobile")
StringGadget(#strMobile, 490, 175, 220, 20, "")
TextGadget(#txtEMail, 490, 200, 100, 20, "E-Mail")
StringGadget(#strEMail, 490, 220, 220, 20, "")
TextGadget(#txtCountry, 260, 260, 80, 20, "Country")
StringGadget(#strCountry, 260, 280, 210, 20, "")
TextGadget(#txtBirthday, 490, 250, 70, 20, "Birthday")
StringGadget(#strBirthday, 570, 250, 120, 20, "")
ButtonGadget(#btnBirthday, 690, 250, 20, 20, "...")
CheckBoxGadget(#chkFamily, 490, 280, 70, 20, "Family")
CheckBoxGadget(#chkFriend, 570, 280, 70, 20, "Friend")
CheckBoxGadget(#chkOther, 640, 280, 70, 20, "Other")
  
;Navigation Buttons
ButtonImageGadget(#btnFirst, 0, 330, 32, 32, ImageID(#imgFirst))
ButtonImageGadget(#btnPrevious, 31, 330, 32, 32, ImageID(#imgPrevious))
ButtonImageGadget(#btnNext, 676, 330, 32, 32, ImageID(#imgNext))
ButtonImageGadget(#btnLast, 708, 330, 32, 32, ImageID(#imgLast))  

ResizeWindow(AddressMain,#PB_Ignore,5,#PB_Ignore,#PB_Ignore)

DisableGadget(#btnBirthday,#True)

OpenAddressDB()
GetTotalRecords()
CurrentRow = 1
Checkrecords()
DisplayRecord()
Mode = "Idle"

Repeat
  
  Event = WaitWindowEvent()
  
  Select EventWindow()
      
    Case PrintControl

      CDPrint::CDPrintEvents(Event)
      
    Default

      Select Event
      
      Case #PB_Event_CloseWindow
        End
      
      Case #PB_Event_Menu
        
        Select EventMenu()
            
          Case #mnuAdd
            
            Mode = "Add"
            ClearDisplay()
            HideIconBarGadget(ToolBarMain,#True)
            HideIconBarGadget(ToolBarEdit,#False)
            DisableGadget(#btnBirthday,#False)
            
          Case #mnuEdit
            
            Mode = "Edit"
            HideIconBarGadget(ToolBarMain,#True)
            HideIconBarGadget(ToolBarEdit,#False)
            DisableGadget(#btnBirthday,#False)
            
          Case #mnuDelete
            
            DeleteContact()
            GetTotalRecords()
            CurrentRow = 1
            CheckRecords()
            DisplayRecord()     
                
          Case #mnuSearch
            
            SearchCriteria = SearchFrm::Open()
            GetTotalRecords()
            CurrentRow = 1
            CheckRecords()
            DisplayRecord()
            
          Case #mnuPageSetup
          
            PageSetup::Open()

          Case #mnuPrintBook
            
            CDPrint::Preview.i = 0 ; Set to print document
            If PrintAddresses() = 0
              MessageRequester("Address Book","Failed To Print",#PB_MessageRequester_Ok |#PB_MessageRequester_Info)
            EndIf           
          
          Case #mnuPreviewBook
          
            CDPrint::Preview.i = 1 ; Set to preview document
            If PrintAddresses() = 1
              CDPrint::ShowPage(0)
            EndIf
          
          Case #mnuExit
          
            End
          
        EndSelect
      
      Case #PB_Event_Gadget
        
        Select EventGadget()
            
          Case #btnBirthday
            
            If CBirthDay > 0
              SelectDate::BirthDay = CBirthDay
              SelectDate::BirthMonth = CBirthMonth
              SelectDate::BirthYear = CBirthYear
            Else
              SelectDate::BirthDay = 1
              SelectDate::BirthMonth = 1
              SelectDate::BirthYear = 2000
            EndIf
            SelectDate::Open()
            CBirthDay = SelectDate::BirthDay
            CBirthMonth = SelectDate::BirthMonth
            CBirthYear = SelectDate::BirthYear
            SetGadgetText(#strBirthday,Str(CBirthDay) + " " + SelectDate::NumberToMonth(CBirthMonth) + " " + Str(CBirthYear))

            
          Case ToolBarMain
            
            Select EventData()
                
              Case 0
                
                Mode = "Add"
                ClearDisplay() 
                HideIconBarGadget(ToolBarMain,#True)
                HideIconBarGadget(ToolBarEdit,#False)
                DisableGadget(#btnBirthday,#False)
                
              Case 1
                
                Mode = "Edit"
                HideIconBarGadget(ToolBarMain,#True)
                HideIconBarGadget(ToolBarEdit,#False)
                DisableGadget(#btnBirthday,#False)
                
              Case 2
                
                DeleteContact()
                GetTotalRecords()
                CurrentRow = 1
                CheckRecords()
                DisplayRecord()                
                
              Case 3
                
                SearchCriteria = SearchFrm::Open()
                GetTotalRecords()
                CurrentRow = 1
                CheckRecords()
                DisplayRecord()
                
              Case 4
                
                CDPrint::Preview.i = 1 ; Set to preview document
                If PrintAddresses() = 1
                  CDPrint::ShowPage(0) 
                EndIf
                
              Case 5
                
                End
                
              Case 6
                
                HelpViewer::Open("")
                
            EndSelect
            
            Case ToolBaredit
            
            Select EventData()          
            
              Case 0

                If Mode = "Add"
                  AddContact()
                EndIf
                SaveContact()
                HideIconBarGadget(ToolBarMain,#False)
                HideIconBarGadget(ToolBarEdit,#True)
                DisableGadget(#btnBirthday,#True) 
                Mode = "Idle"
                
              Case 1
                
                HideIconBarGadget(ToolBarMain,#False)
                HideIconBarGadget(ToolBarEdit,#True)
                DisableGadget(#btnBirthday,#True) 
                Mode = "Idle"
                CheckRecords()
                DisplayRecord()
                
            EndSelect
            
          Case #btnFirst
          
            CurrentRow = 1
            CheckRecords()
            DisplayRecord()
            
          Case #btnPrevious
          
            If CurrentRow > 1
              CurrentRow = CurrentRow - 1
              CheckRecords()
              DisplayRecord()
            EndIf          
          
          Case #btnNext

            If CurrentRow < TotalRows
              CurrentRow = CurrentRow + 1
              CheckRecords()
              DisplayRecord()
            EndIf
          
          Case #btnLast
                    
            CurrentRow = TotalRows
            CheckRecords()
            DisplayRecord()
          
          Case #btnAssign
          
            AddThumb()
          
        EndSelect   ;EventGadget()       
          
      EndSelect ;Event

  EndSelect  ;Eventwindow()
  
 ForEver  
 
 DataSection
  FirstPhoto: 
    IncludeBinary "Resources\Resultset_first.png"
  PreviousPhoto: 
    IncludeBinary "Resources\Resultset_previous.png"
  NextPhoto: 
    IncludeBinary "Resources\Resultset_next.png"
  LastPhoto: 
    IncludeBinary "Resources\Resultset_last.png"
  ToolBarAdd:
    IncludeBinary "Resources\Add.png" 
  ToolBarEdit:
    IncludeBinary "Resources\Edit.png"   
  ToolBarDelete:
    IncludeBinary "Resources\Delete.png"   
  ToolBarSearch:
    IncludeBinary "Resources\Search.png"
  ToolBarPreferences:
    IncludeBinary "Resources\Preferences.png" 
  ToolBarExit:
    IncludeBinary "Resources\Exit.png"  
  ToolBarOk:
    IncludeBinary "Resources\Ok.png" 
  ToolBarCancel:
    IncludeBinary "Resources\Cancel.png"  
  ToolBarHelp:
    IncludeBinary "Resources\Help.png"    
  ToolBarPrint:
    IncludeBinary "Resources\Print.png"     
  EndDataSection
; IDE Options = PureBasic 5.60 Beta 1 (Windows - x64)
; CursorPosition = 667
; FirstLine = 437
; Folding = Aw
; EnableXP