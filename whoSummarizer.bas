Attribute VB_Name = "Module1"
Public wapp As Object
Public wdoc As Object
Public counter As Integer
 
Sub main()
Range(Cells(2, 1), Cells(Range("a2").CurrentRegion.Rows.Count, Range("a2").CurrentRegion.Columns.Count)).ClearContents
Set wdoc = CreateObject("Word.Document")
Set wapp = CreateObject("Word.Application")

'select workning directory
Application.FileDialog(msoFileDialogFolderPicker).title = "Please Select A Folder To Save The Results"
folder = Application.FileDialog(msoFileDialogFolderPicker).Show


If folder <> 0 Then
    folder = Application.FileDialog(msoFileDialogFolderPicker).SelectedItems(1) + "\"
Else

MsgBox "You have not selected any folders please select a folder to save outputs"

If counter > 0 Then
GoTo ended


Else
counter = counter + 1
Call main
End If
End If

'make word document
Set wdoc = wapp.Documents.Add
Call wordHeadingFormats
wapp.Visible = True


cont = "y"

While cont = "y"
'select which journals?
journal = InputBox("Please Enter Journal Number:" + Chr(10) + "1.The Lancet" + Chr(10) + "2.Lancet Global Health" + Chr(10) + "3.Lancet Digital Health" + Chr(10) + "4.Lancet Diabetes" + Chr(10) + _
"5.Lancet Public Health" + Chr(10) + "6.New England" + Chr(10) _
+ "7.Jama" + Chr(10) + "8.Journal of American Epidemiology" + Chr(10) + "9.Journal of European Epidemiology(Needs Access)" _
+ Chr(10) + "10.Plos One" + Chr(10) + "11.Nature(Needs Access)" + Chr(10) + "12.Diabetes Care")
website = InputBox("Please Enter Link of the current issue? (leave blank if unsure)")

'select case
Select Case journal

Case 1, 2, 3, 4, 5
    If website = Empty Then
    
      If journal = 1 Then
     website = "https://www.thelancet.com/journals/lancet/issue/current"
      
      ElseIf journal = 2 Then
      website = "https://www.thelancet.com/journals/langlo/issue/current"
      
       ElseIf journal = 3 Then
      website = "https://www.thelancet.com/journals/landig/issue/current"
      
       ElseIf journal = 4 Then
      website = "https://www.thelancet.com/journals/landia/issue/current"
      
       ElseIf journal = 5 Then
      website = "https://www.thelancet.com/journals/lanpub/issue/current"
     
    End If
    End If
    Call lancet(website)


Case 6
    Call nejm(website)


Case 7
    Call jama(website)

Case 8

    Call jae(website)

Case 9, 11
MsgBox "Needs subscription, please contact system designer!"
  '  Call jee(website)

Case 10

    Call plos(website)

Case 11
MsgBox "Needs subscription, please contact system designer!"
   ' Call nature(website)

Case 12
    Call diabCare(website)
Case ""
cont = "n"
    
End Select
prog.Hide
If cont = "y" Then
 cont = InputBox("Done! Do you want to continue with another journal? y/n")
End If
 Wend
wdoc.Paragraphs.Add
rws = Range("a1").CurrentRegion.Rows.Count
For i = 2 To rws
    ps = wdoc.Paragraphs.Count
       cls = Range("a" + CStr(i)).End(xlToRight).Column
       If cls > 3 Then
    'add journal
    If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
        
        wdoc.Paragraphs(ps).Range.InsertAfter Cells(i, 1).Value
        wdoc.Paragraphs(ps).Style = "Heading 1"
        wdoc.Paragraphs.Add
        ps = wdoc.Paragraphs.Count
    End If
    'add title
    wdoc.Paragraphs.Add
    wdoc.Paragraphs(ps + 1).Range.InsertAfter Cells(i, 3).Value
    wdoc.Paragraphs(ps + 1).Range.Hyperlinks.Add Anchor:=wdoc.Paragraphs(ps + 1).Range, Address:=Cells(i, 2).Value
    wdoc.Paragraphs(ps + 1).Style = "Heading 2"
    wdoc.Paragraphs.Add

    'add details

    For k = 4 To cls
        ps = wdoc.Paragraphs.Count
        wdoc.Paragraphs(ps).Range.InsertAfter Cells(i, k)
        wdoc.Paragraphs(ps).Range.InsertParagraphAfter
    Next
    End If
Next

wdoc.TablesOfContents.Add wdoc.Paragraphs(1).Range, UseHyperlinks:=True, UpperHeadingLevel:=1, LowerHeadingLevel:=2, UseHeadingStyles:=True


wdoc.SaveAs2 folder + Replace(Date, "/", "-") + "-journalReport"
wdoc.SaveAs2 folder + Replace(Date, "/", "-") + "-journalReport.pdf", FileFormat:=17
 
wdoc.Close
Set wapp = Nothing
'heading and body text of word document
    MsgBox "The Files were created successfully in the selected folder!"
ended:
    ThisWorkbook.Save
    Shell "C:\WINDOWS\explorer.exe " & folder, vbNormalFocus
    ThisWorkbook.Close
   



    




End Sub


Sub jama(website)
    ' Date:16/09/2021
    ' Author: Yosef Farzi
    ' QA: Shahedeh Seyfi
    
    ' variable definitions
    Set chbrowser = New Selenium.ChromeDriver
    
    Dim divs        As Selenium.WebElements
    Dim articles    As Selenium.WebElements
    Dim div         As Selenium.WebElement
    Dim article     As Selenium.WebElement
    Dim parags      As Selenium.WebElements
    Dim parag       As Selenium.WebElement
    
    'Initialize progress bar
    prog.Show False
    prog.progressBar.Width = 20
    prog.title.Caption = "Finding Original Articles"
    
    'Start and load main page of the journal
    If website = Empty Then
        website = "https://jamanetwork.com/journals/jama/currentissue"
    End If
    chbrowser.Start
    Call chbrowser.Get(Url:=website, timeout:=7000, Raise:=0)
    brws = Range("b1").CurrentRegion.Rows.Count
    rws = brws
    'find only origianl articles
    Set divs = chbrowser.FindElementsByClass("group--original-investigation")
    For Each div In divs
        If div.FindElementByClass("article-type-group").Text = "Original Investigation" Then
            Set articles = div.FindElementsByClass("article")
            
            ' Record address and journal of found articles
            For Each article In articles
                rws = rws + 1
                Cells(rws, 2).Value = article.FindElementByTag("h3").FindElementByTag("a").Attribute("href")
                Cells(rws, 3).Value = article.FindElementByTag("h3").FindElementByTag("a").Text
                Cells(rws, 1).Value = chbrowser.FindElementByClass("network").Text
            Next
            
            'Update progress bar
            For i = brws + 1 To rws
                percentC = ((i - brws) / (rws - brws)) * 200
                prog.progressBar.Width = percentC
                prog.title.Caption = "Working On Article Number :" + CStr(i - 1)
                VBA.DoEvents
                
                'Open each article to get main point and save to excel
                Call chbrowser.Get(Url:=Cells(i, 2).Value, timeout:=10000, Raise:=0)
                Set parags = chbrowser.FindElementByClass("article-full-text").FindElementsByTag("p")
                c = 1
                For Each parag In parags
                    Cells(i, 3 + c).Value = parag.Text
                    c = c + 1
                Next
                
            Next
        
        End If
    Next

    prog.Hide

End Sub



Sub jae(website)
    ' Date:16/09/2021
    ' Author: Yosef Farzi, Shahedeh Seyfi
    ' QA:
    
    ' variable definitions
    Set chbrowser = New Selenium.ChromeDriver
    
    Dim divs        As Selenium.WebElements
    Dim articles    As Selenium.WebElements
    Dim div         As Selenium.WebElement
    Dim article     As Selenium.WebElement
    Dim parags      As Selenium.WebElements
    Dim parag       As Selenium.WebElement
    
    'Initialize progress bar
    prog.Show False
    prog.progressBar.Width = 20
    prog.title.Caption = "Finding Original Articles"
    
    'Start and load main page of the journal
    If website = Empty Then
    website = "https://academic.oup.com/aje/issue/190/9#1266799-6329297"
    End If
    
    chbrowser.Start
    Call chbrowser.Get(Url:=website, timeout:=20000, Raise:=0)
    brws = Range("b1").CurrentRegion.Rows.Count
    rws = brws
    'find only origianl articles
    Set divs = chbrowser.FindElementById("resourceTypeList-OUP_Issue").FindElementsByTag("section")
    For Each div In divs

        If div.FindElementByTag("h4").Text = "ORIGINAL CONTRIBUTIONS" Then
            Set articles = div.FindElementsByClass("al-article-items")
            
           '  Record address and journal of found articles
          For Each article In articles
                rws = rws + 1
               Cells(rws, 1).Value = chbrowser.FindElementById("logo-AmericanJournalofEpidemiology").Attribute("alt")
                Cells(rws, 2).Value = article.FindElementByTag("h5").FindElementByTag("a").Attribute("href")
                Cells(rws, 3).Value = article.FindElementByTag("h5").FindElementByTag("a").Text
            Next
            
            'Update progress bar
            For i = brws + 1 To rws
                percentC = ((i - brws) / (rws - brws)) * 200
                prog.progressBar.Width = percentC
               prog.title.Caption = "Working On Article Number :" + CStr(i - 1)
                VBA.DoEvents
                
                'Open each article to get main point and save to excel
                Call chbrowser.Get(Url:=Cells(i, 2).Value, timeout:=7000, Raise:=0)
                If chbrowser.FindElementsByClass("abstract-title").Count > 0 Then
                   Set parags = chbrowser.FindElementByClass("abstract").FindElementsByTag("p")
                   c = 1
                    For Each parag In parags
                    Cells(i, 3 + c).Value = parag.Text
                    c = c + 1
                
                       Next
               End If
                    Next
              Exit For
       End If
 
    Next
    
End Sub



Sub jee(website)
    ' Date:16/09/2021
    ' Author: Yosef Farzi
    ' QA: Shahedeh Seyfi
    
    ' variable definitions
    Set chbrowser = New Selenium.ChromeDriver
    
    Dim divs        As Selenium.WebElements
    Dim articles    As Selenium.WebElements
    Dim div         As Selenium.WebElement
    Dim article     As Selenium.WebElement
    Dim parags      As Selenium.WebElements
    Dim parag       As Selenium.WebElement
    
    'Initialize progress bar
    prog.Show False
    prog.progressBar.Width = 20
    prog.title.Caption = "Finding Original Articles"
    
    'Start and load main page of the journal
    If website = Empty Then
    website = "https://link.springer.com/journal/10654/volumes-and-issues/36-8"
    End If
    chbrowser.Start
    Call chbrowser.Get(Url:=website, timeout:=10000, Raise:=0)
    brws = Range("b1").CurrentRegion.Rows.Count
    rws = brws
    'find only origianl articles
    Set divs = chbrowser.FindElementByClass("app-volumes-and-issues__article")
    For Each div In divs

        If div.FindElementByTag("h4").Text = "ORIGINAL CONTRIBUTIONS" Then
            Set articles = div.FindElementsByClass("c-card__title")
            
            ' Record address and journal of found articles
            For Each article In articles
                rws = rws + 1
                Cells(rws, 1).Value = chbrowser.FindElementById("journalTitle").FindElementByTag("a").Text
                Cells(rws, 2).Value = article.FindElementByTag("h3").FindElementByTag("a").Attribute("href")
                Cells(rws, 3).Value = article.FindElementByTag("h3").FindElementByTag("a").Text
            Next
            
            'Update progress bar
            For i = brws + 1 To rws
                percentC = ((i - brws) / (rws - brws)) * 200
                prog.progressBar.Width = percentC
                prog.title.Caption = "Working On Article Number :" + CStr(i - 1)
                VBA.DoEvents
                
                'Open each article to get main point and save to excel
               Call chbrowser.Get(Url:=Cells(i, 2).Value, timeout:=7000, Raise:=0)
                If chbrowser.FindElementsByClass("abstract-title").Count > 0 Then
                    Set parags = chbrowser.FindElementByClass("abstract").FindElementsByTag("p")
                    c = 1
                    For Each parag In parags
                    Cells(i, 3 + c).Value = parag.Text
                    c = c + 1
                
                        Next
                End If
                    Next
           Exit For
         End If
    Next
    
End Sub
Sub lancet(website)
    ' Date:16/09/2021
    ' Author: Yosef Farzi
    ' QA: Shahedeh Seyfi
    
    ' variable definitions
    Set chbrowser = New Selenium.ChromeDriver
    
    Dim divs        As Selenium.WebElements
    Dim articles    As Selenium.WebElements
    Dim div         As Selenium.WebElement
    Dim article     As Selenium.WebElement
    Dim parags      As Selenium.WebElements
    Dim heads       As Selenium.WebElements
    Dim parag       As Selenium.WebElement
    
    'Initialize progress bar
    prog.Show False
    prog.progressBar.Width = 20
    prog.title.Caption = "Finding Original Articles"
    
    'Start and load main page of the journal
    chbrowser.Start
    Call chbrowser.Get(Url:=website, timeout:=30000, Raise:=0)
    brws = Range("b1").CurrentRegion.Rows.Count
    rws = brws
    'find only origianl articles
    Set divs = chbrowser.FindElementByClass("table-of-content__section").FindElementsByTag("section")
    For Each div In divs

        If div.FindElementByTag("h2").Text = "ARTICLES" Then
            Set articles = div.FindElementsByClass("toc__item__body")
            
            ' Record address and journal of found articles
            For Each article In articles
                rws = rws + 1
                Cells(rws, 1).Value = chbrowser.FindElementByClass("journal-logos").FindElementByTag("img").Attribute("alt")
                Cells(rws, 2).Value = article.FindElementByClass("toc__item__title").FindElementByTag("a").Attribute("href")
                Cells(rws, 3).Value = article.FindElementByClass("toc__item__title").FindElementByTag("a").Text
            Next
            
            'Update progress bar
            For i = brws + 1 To rws
                percentC = ((i - brws) / (rws - brws)) * 200
                prog.progressBar.Width = percentC
                prog.title.Caption = "Working On Article Number :" + CStr(i - brws)
                VBA.DoEvents
                
                'Open each article to get main point and save to excel
                Call chbrowser.Get(Url:=Cells(i, 2).Value, timeout:=15000, Raise:=0)
                If chbrowser.FindElementsByTag("section").Count > 0 Then
                    Set parags = chbrowser.FindElementByClass("section-paragraph").FindElementsByClass("section-paragraph")
                    Set heads = chbrowser.FindElementByClass("section-paragraph").FindElementsByTag("h3")
                    c = 1
                    For j = 1 To parags.Count
                    Cells(i, 3 + c).Value = heads(j).Text
                    Cells(i, 3 + c + 1).Value = parags(j).Text
                    c = c + 2
                
                        Next
                End If
                    Next
           Exit For
        End If
     
    Next
    
End Sub


Sub diabCare(website)
    ' Date:16/09/2021
    ' Author: Yosef Farzi
    ' QA: Shahedeh Seyfi
    
    ' variable definitions
    Set chbrowser = New Selenium.ChromeDriver
    
    Dim divs        As Selenium.WebElements
    Dim articles    As Selenium.WebElements
    Dim div         As Selenium.WebElement
    Dim article     As Selenium.WebElement
    Dim parags      As Selenium.WebElements
    Dim heads       As Selenium.WebElements
    Dim parag       As Selenium.WebElement
    
    'Initialize progress bar
    prog.Show False
    prog.progressBar.Width = 20
    prog.title.Caption = "Finding Original Articles"
    
    'Start and load main page of the journal
    If website = Empty Then
    website = "https://care.diabetesjournals.org/content/44/9?current-issue=y"
    End If
    chbrowser.Start
    Call chbrowser.Get(Url:=website, timeout:=15000, Raise:=0)
    brws = Range("b1").CurrentRegion.Rows.Count
    rws = brws
    'find only origianl articles
    Set divs = chbrowser.FindElementsByClass("issue-toc-section")
    For Each div In divs

        If div.FindElementByTag("h2").Text <> "Addenda" And div.FindElementByTag("h2").Text <> "Commentaries" And div.FindElementByTag("h2").Text <> "Issues and Events" And div.FindElementByTag("h2").Text <> "e-Letters" Then
            Set articles = div.FindElementsByClass("highwire-cite")
            
            ' Record address and journal of found articles
            For Each article In articles
                rws = rws + 1
                Cells(rws, 1).Value = chbrowser.FindElementByClass("logo-img").FindElementByTag("img").Attribute("alt")
                Cells(rws, 2).Value = article.FindElementByTag("a").Attribute("href")
                Cells(rws, 3).Value = article.FindElementByTag("a").Text
            Next
                    

        
        End If
      Next
            'Update progress bar
            For i = brws + 1 To rws
                percentC = ((i - brws) / (rws - brws)) * 200
                prog.progressBar.Width = percentC
                prog.title.Caption = "Working On Article Number :" + CStr(i - brws)
                VBA.DoEvents
                
                'Open each article to get main point and save to excel
                Call chbrowser.Get(Url:=Cells(i, 2).Value, timeout:=25000, Raise:=0)
                If chbrowser.FindElementsByClass("abstract").Count > 0 Then
                   
                    Set parags = chbrowser.FindElementByClass("abstract").FindElementsByClass("subsection")
                    c = 1
                    For j = 1 To parags.Count
                    Cells(i, 3 + c).Value = parags(j).Text
                    
                    c = c + 1
                
                        Next
                End If
   
     
    Next
    
End Sub

Sub plos(website)
    ' Date:16/09/2021
    ' Author: Yosef Farzi
    ' QA: Shahedeh Seyfi
    
    ' variable definitions
    Set chbrowser = New Selenium.ChromeDriver
    
    Dim divs        As Selenium.WebElements
    Dim articles    As Selenium.WebElements
    Dim div         As Selenium.WebElement
    Dim article     As Selenium.WebElement
    Dim parags      As Selenium.WebElements
    Dim heads       As Selenium.WebElements
    Dim parag       As Selenium.WebElement
    
    'Initialize progress bar
    prog.Show False
    prog.progressBar.Width = 20
    prog.title.Caption = "Finding Original Articles"
    
    'Start and load main page of the jurnal
    If website = Empty Then
    website = "https://journals.plos.org/plosmedicine/issue"
    End If
    chbrowser.Start
    Call chbrowser.Get(Url:=website, timeout:=20000, Raise:=0)
    brws = Range("b1").CurrentRegion.Rows.Count
    rws = brws
    'find only origianl articles
    Set divs = chbrowser.FindElementsByClass("section")
    For Each div In divs

        If div.FindElementByTag("h2").Text = "Research Articles" Then
            Set articles = div.FindElementsByClass("item--article-title")
            
            ' Record address and journal of found articles
            For Each article In articles
                rws = rws + 1
                Cells(rws, 1).Value = chbrowser.FindElementByClass("logo").Text
                Cells(rws, 2).Value = article.FindElementByTag("a").Attribute("href")
                Cells(rws, 3).Value = article.FindElementByTag("a").Text
            Next
            
            'Update progress bar
            For i = brws + 1 To rws
                percentC = ((i - brws) / (rws - brws)) * 200
                prog.progressBar.Width = percentC
                prog.title.Caption = "Working On Article Number :" + CStr(i - brws)
                VBA.DoEvents
                
                'Open each article to get main point and save to excel
                Call chbrowser.Get(Url:=Cells(i, 2).Value, timeout:=15000, Raise:=0)
                If chbrowser.FindElementsByTag("section").Count > 0 Then
                    Set parags = chbrowser.FindElementByClass("article-content").FindElementsByClass("abstract-content")
                    Set heads = chbrowser.FindElementByClass("abstract-content").FindElementsByTag("h3")
                    c = 1
                    For j = 1 To parags.Count
                    Cells(i, 3 + c).Value = heads(j).Text
                    Cells(i, 3 + c + 1).Value = parags(j).Text
                    c = c + 2
                
                        Next
                End If
                    Next
           Exit For
        End If
     
    Next
    
End Sub


Sub nejm(website)
    ' Date:16/09/2021
    ' Author: Yosef Farzi
    ' QA: Shahedeh Seyfi
    
    ' variable definitions
    Set chbrowser = New Selenium.ChromeDriver
    Dim sections    As Selenium.WebElements
    Dim divs        As Selenium.WebElements
    Dim articles    As Selenium.WebElements
    Dim div         As Selenium.WebElement
    Dim article     As Selenium.WebElement
    Dim parags      As Selenium.WebElements
    Dim heads       As Selenium.WebElements
    Dim parag       As Selenium.WebElement
    
    'Initialize progress bar
    prog.Show False
    prog.progressBar.Width = 20
    prog.title.Caption = "Finding Original Articles"
    
    'Start and load main page of the journal
    If website = Empty Then
    website = "https://www.nejm.org/toc/nejm/medical-journal"
    End If
    
    chbrowser.Start
    Call chbrowser.Get(Url:=website, timeout:=20000, Raise:=0)
    brws = Range("b1").CurrentRegion.Rows.Count
    rws = brws
    'find only origianl articles
    Set divs = chbrowser.FindElementByClass("o-col--primary").FindElementsByClass("o-results")
    Set sections = chbrowser.FindElementByClass("o-col--primary").FindElementsByTag("h2")
    For k = 1 To sections.Count
        If sections(k).Text = "ORIGINAL ARTICLES" Then
            Set div = divs(k)
            Exit For
        End If
    Next

    

    Set articles = div.FindElementsByClass("m-teaser-item__link")
    
    ' Record address and journal of found articles
    For Each article In articles
        rws = rws + 1
        Cells(rws, 1).Value = "New England Journal of Medicine"
        Cells(rws, 2).Value = article.Attribute("href")
        Cells(rws, 3).Value = article.Text
    Next
    
    'Update progress bar
    For i = brws + 1 To rws
        percentC = ((i - brws) / (rws - brws)) * 200
        prog.progressBar.Width = percentC
        prog.title.Caption = "Working On Article Number :" + CStr(i - brws)
        VBA.DoEvents
        
        'Open each article to get main point and save to excel
        Call chbrowser.Get(Url:=Cells(i, 2).Value, timeout:=15000, Raise:=0)
        If chbrowser.FindElementsById("article_Abstract").Count > 0 Then
            Set parags = chbrowser.FindElementById("article_Abstract").FindElementsByTag("p")
            Set heads = chbrowser.FindElementById("article_Abstract").FindElementsByTag("h2")
            c = 1
            For j = 1 To parags.Count
            Cells(i, 3 + c).Value = heads(j).Text
            Cells(i, 3 + c + 1).Value = parags(j).Text
            c = c + 2
        
            Next
        End If
    Next
        
       
     
  
    
End Sub

Sub nature(website)
    ' Date:16/09/2021
    ' Author: Yosef Farzi
    ' QA: Shahedeh Seyfi
    
    ' variable definitions
    Set chbrowser = New Selenium.ChromeDriver
    Dim sections    As Selenium.WebElements
    Dim divs        As Selenium.WebElements
    Dim articles    As Selenium.WebElements
    Dim div         As Selenium.WebElement
    Dim article     As Selenium.WebElement
    Dim parags      As Selenium.WebElements
    Dim heads       As Selenium.WebElements
    Dim parag       As Selenium.WebElement
    
    'Initialize progress bar
    prog.Show False
    prog.progressBar.Width = 20
    prog.title.Caption = "Finding Original Articles"
    
    'Start and load main page of the journal
    If website = Empty Then
    website = "https://www.nature.com/nature/current-issue"
    End If
    chbrowser.Start
    Call chbrowser.Get(Url:=website, timeout:=20000, Raise:=0)
    brws = Range("b1").CurrentRegion.Rows.Count
    rws = brws
    'find only origianl articles
  

    Set articles = chbrowser.FindElementById("ThisWeek-content").FindElementsByTag("article")
    
    ' Record address and journal of found articles
    For Each article In articles
    If article.FindElementByTag("span").Text = "Research Highlight" Then
        rws = rws + 1
        Cells(rws, 1).Value = chbrowser.FindElementByClass("c-header__logo-container").FindElementByTag("img").Attribute("alt")
        Cells(rws, 2).Value = article.FindElementByTag("h3").FindElementByTag("a").Attribute("href")
        Cells(rws, 3).Value = article.FindElementByTag("h3").FindElementByTag("a").Text
        
    End If
    Next
    MsgBox ("needs access!")
    Exit Sub
    'Update progress bar
    For i = brws + 1 To rws
        percentC = ((i - brws) / (rws - brws)) * 200
        prog.progressBar.Width = percentC
        prog.title.Caption = "Working On Article Number :" + CStr(i - brws)
        VBA.DoEvents
        
        'Open each article to get main point and save to excel
        Call chbrowser.Get(Url:=Cells(i, 2).Value, timeout:=15000, Raise:=0)
        If chbrowser.FindElementsById("article_Abstract").Count > 0 Then
            Set parags = chbrowser.FindElementById("article_Abstract").FindElementsByTag("p")
            Set heads = chbrowser.FindElementById("article_Abstract").FindElementsByTag("h2")
            c = 1
            For j = 1 To parags.Count
            Cells(i, 3 + c).Value = heads(j).Text
            Cells(i, 3 + c + 1).Value = parags(j).Text
            c = c + 2
        
            Next
        End If
    Next
        
       
     
  
    
End Sub
Sub wordHeadingFormats()



    With wdoc.Styles("Heading 1").Font
        .Name = "Arial Rounded MT Bold"
        .Size = 20
        
        .Color = -738148353
        
    End With
     wdoc.Styles("Heading 1").NoSpaceBetweenParagraphsOfSameStyle = _
        False
    With wdoc.Styles("Heading 1")
        .AutomaticallyUpdate = False
        .BaseStyle = "Normal"
        .NextParagraphStyle = "Normal"
    End With
    wdoc.Styles("Heading 1").ParagraphFormat.PageBreakBefore = True
    wdoc.Styles("Heading 2").ParagraphFormat.PageBreakBefore = True
    With wdoc.Styles("Heading 2").Font
        .Name = "Arial Rounded MT Bold"
        .Size = 16
        .Color = -738148353
    End With
    With wdoc.Styles("Heading 2")
        .AutomaticallyUpdate = False
        .BaseStyle = "Normal"
        .NextParagraphStyle = "Normal"
    End With

End Sub

