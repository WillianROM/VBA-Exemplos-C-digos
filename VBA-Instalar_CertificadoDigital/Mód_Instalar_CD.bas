Attribute VB_Name = "Mód_ImportCD"

'Nas referências tem que deixar marcado a opção UIAutomationClient

Dim oAutomation As New CUIAutomation

Sub ImportacaoDeCertificadoDigital()
    'Localizar a janela: Assistente para Importação de Certificados
    Dim janela As UIAutomationClient.IUIAutomationElement
    Set janela = WalkEnabledElements(oAutomation.GetRootElement, "Assistente para Importação de Certificados")
 
    'Localizar o botão avançar
    Dim btnAvancar As UIAutomationClient.IUIAutomationElement
    Set btnAvancar = janela.FindFirst(TreeScope_Children, PropCondition(oAutomation, "Avançar", "Name"))
    
    'Clicar no botão Avançar duas vezes seguidas
    Dim btnAvancarClick As UIAutomationClient.IUIAutomationInvokePattern
    Set btnAvancarClick = btnAvancar.GetCurrentPattern(UIAutomationClient.UIA_InvokePatternId)
    btnAvancarClick.Invoke
    btnAvancarClick.Invoke
    
    'Localizar a "Caixa de diálogo"
    Dim caixaDeDialogo As UIAutomationClient.IUIAutomationElement
    Set caixaDeDialogo = janela.FindFirst(TreeScope_Children, PropCondition(oAutomation, "Win32PropSheetPageHost", "ClsName"))
    
    'Localizar o campo Senha pelo LocalizedControlType Senha
    Dim txtSenha As UIAutomationClient.IUIAutomationElement
    Set txtSenha = caixaDeDialogo.FindFirst(TreeScope_Children, PropCondition(oAutomation, "editar", "LoczCon"))

    'Informar a senha no campo Senha
    Dim oPattern As UIAutomationClient.IUIAutomationLegacyIAccessiblePattern
    Set oPattern = txtSenha.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
    oPattern.SetValue ("12345678")

    'Clicar no botão Avançar duas vezes seguidas
    btnAvancarClick.Invoke
    btnAvancarClick.Invoke
    
    'Localizar o botão Concluir
    Dim btnConcluir As UIAutomationClient.IUIAutomationElement
    Set btnConcluir = janela.FindFirst(TreeScope_Children, PropCondition(oAutomation, "Concluir", "Name"))
    
    'Clicar no botão Concluir
    Dim btnConcluirClick As UIAutomationClient.IUIAutomationInvokePattern
    Set btnConcluirClick = btnConcluir.GetCurrentPattern(UIAutomationClient.UIA_InvokePatternId)
    btnConcluirClick.Invoke
    
    'Aguardar 1 segundo
    Application.Wait (Now + TimeValue("00:00:01"))
    
    'Localizar a janela: Assistente para Importação de Certificados
    Dim janela2 As UIAutomationClient.IUIAutomationElement
    Set janela2 = WalkEnabledElements(oAutomation.GetRootElement, "Assistente para Importação de Certificados")
   
    'Localizar o botão OK
    Dim btnOK As UIAutomationClient.IUIAutomationElement
    Set btnOK = janela2.FindFirst(TreeScope_Children, PropCondition(oAutomation, "OK", "Name"))
    
    'Clicar no botão Concluir
    Dim btnOKClick As UIAutomationClient.IUIAutomationInvokePattern
    Set btnOKClick = btnOK.GetCurrentPattern(UIAutomationClient.UIA_InvokePatternId)
    btnOKClick.Invoke
    

End Sub


Function WalkEnabledElements(element As UIAutomationClient.IUIAutomationElement, strWIndowName As String) As UIAutomationClient.IUIAutomationElement

    Dim walker As UIAutomationClient.IUIAutomationTreeWalker
    
    Set walker = oAutomation.ControlViewWalker
    Set element = walker.GetFirstChildElement(element)
    
    Do While Not element Is Nothing
    
        If InStr(1, element.CurrentName, strWIndowName) > 0 Then
            Set WalkEnabledElements = element
            Exit Function
        End If
        Set element = walker.GetNextSiblingElement(element)
    Loop
End Function

Function PropCondition(UIAutomation As CUIAutomation, Requirement As String, IdType As String) As UIAutomationClient.IUIAutomationCondition
    Select Case IdType
        Case "Name":
            Set PropCondition = UIAutomation.CreatePropertyCondition(UIAutomationClient.UIA_NamePropertyId, Requirement)
        Case "AutoID":
            Set PropCondition = UIAutomation.CreatePropertyCondition(UIAutomationClient.UIA_AutomationIdPropertyId, Requirement)
        Case "ClsName":
            Set PropCondition = UIAutomation.CreatePropertyCondition(UIAutomationClient.UIA_ClassNamePropertyId, Requirement)
        Case "LoczCon":
            Set PropCondition = UIAutomation.CreatePropertyCondition(UIAutomationClient.UIA_LocalizedControlTypePropertyId, Requirement)
    End Select
End Function

