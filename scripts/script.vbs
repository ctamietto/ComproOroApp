	'Dim kt as string
		
	' testata spedizione
	Class spedclass
    	Public kt,banco_metalli_id,banco_metalli_desc , data_ddt, numero_ddt
    	'Public titolo_oro_id, titolo_oro_desc
    	Public verga_stimata , titolo_stimato_verga , totale_grammi_rottami , totale_grammi_puro_stimato
    	Public verga_fonderia , titolo_fonderia , totale_grammi_puro_fonderia
    	Public titolo_lab_controsaggio , puro_stimato_lab_controsaggio
    	Public differenza_grammi,differenza_percentuale  
    	Public differenza_grammi_con_saggio,differenza_percentuale_con_saggio
	End Class

	' dettaglio spedizione 
	Class speddetclass
		Public kt,fk,titolo_oro_id,titolo_oro_desc,grammi_lordi,grammi_puro_stimati	
	End Class 

	Class bmclass
    	Public kt,desc
	End Class

	Class toclass
    	Public kt,desc,coefficiente,coefficiente_titolo_stimato
	End Class

	Class errorclass
		Public cod,tipo,field,desc
	End Class

	Public speddict: Set speddict = CreateObject("Scripting.Dictionary")
	
	Public speddetdict: Set speddetdict = CreateObject("Scripting.Dictionary")

	Public currspeddetdict: Set currspeddetdict = CreateObject("Scripting.Dictionary")

	Public bmdict: Set bmdict = CreateObject("Scripting.Dictionary")

	Public todict: Set todict = CreateObject("Scripting.Dictionary")
	
	Public spederrorsdict: Set spederrorsdict = CreateObject("Scripting.Dictionary")

	Sub initialSizeAndPos
		window.moveTo 30, 30
	End Sub 

	Function SpedErrorsStatus
		Dim ES
		ES = False  
		If Not IsNull(spederrorsdict) And Not IsEmpty(spederrorsdict) And spederrorsdict.Count > 0 Then
			ES = True 
		End If 
		SpedErrorsStatus = ES
	End Function
	
	Sub SpedErrorsAdd(errcl)
		spederrorsdict.Add errcl.field,errcl
		'alert(errcl.desc)
	End Sub 

	Sub SpedErrorsCleared
		spederrorsdict.RemoveAll()	
	End Sub

	Sub SpedDetailValidate()
  		Dim existErrors
		existErrors = False 
		SpedErrorsCleared()
			
		Dim sped: Set sped = new spedclass
		Dim isNewRecord
		isNewRecord = False
				
		Call SpedDetailGet(sped,isNewRecord,False)
		calcSped(sped) 
		SpedDetailDisplay(sped)
    End Sub
    
	Sub SpedDetailCopy(ByRef spedDest,spedSrc)
	
		spedDest.kt = spedSrc.kt		

		spedDest.banco_metalli_id = spedSrc.banco_metalli_id
		
		spedDest.numero_ddt = spedSrc.numero_ddt
		
		spedDest.data_ddt = spedSrc.data_ddt
		
		'spedDest.titolo_oro_id = spedSrc.titolo_oro_id
		
		spedDest.verga_stimata = spedSrc.verga_stimata
		
		spedDest.titolo_stimato_verga = spedSrc.titolo_stimato_verga
		
		spedDest.totale_grammi_rottami = spedSrc.totale_grammi_rottami
		
		spedDest.totale_grammi_puro_stimato = spedSrc.totale_grammi_puro_stimato
		
		spedDest.verga_fonderia = spedSrc.verga_fonderia
		
		spedDest.titolo_fonderia = spedSrc.titolo_fonderia
		
		spedDest.totale_grammi_puro_fonderia = spedSrc.totale_grammi_puro_fonderia
		
		spedDest.titolo_lab_controsaggio = spedSrc.titolo_lab_controsaggio
		
		spedDest.puro_stimato_lab_controsaggio = spedSrc.puro_stimato_lab_controsaggio
				
		spedDest.differenza_grammi = spedSrc.differenza_grammi
		spedDest.differenza_percentuale = spedSrc.differenza_percentuale
		spedDest.differenza_grammi_con_saggio = spedSrc.differenza_grammi_con_saggio
		spedDest.differenza_percentuale_con_saggio = spedSrc.differenza_percentuale_con_saggio
	
    End Sub
    

	Sub SpedDetailGet(ByRef sped,ByRef isNewRecord,genKT)
		'MsgBox "init SpedDetailGet"
		SpedErrorsCleared()
		isNewRecord = False 
		anyErrors = False 
		
		Set kt = document.getElementById( "kt" )
		If (IsNull(kt.value) Or IsEmpty(kt.value)) Then
			sped.kt = ""
		Else 
			sped.kt = kt.value
		End If		
		
		Set banco_metalli = document.getElementById( "banco_metalli" )

		If (IsNull(banco_metalli.value) Or IsEmpty(banco_metalli.value)) Then
			sped.banco_metalli_id = ""
		Else 
			sped.banco_metalli_id  = banco_metalli.value
		End If		
		
		If (sped.banco_metalli_id = "") Then
			Dim errclBMI: Set errclBMI = New errorclass
			errclBMI.cod   = "000001"
			errclBMI.tipo  = "REQUIRED"
			errclBMI.field = "banco_metalli_id"
			errclBMI.desc  = "IMPOSTARE BANCO METALLI"
			SpedErrorsAdd(errclBMI)
			anyErrors = True  
		End If
				
		Set numero_ddt = document.getElementById( "numero_ddt" )

		If (IsNull(numero_ddt.value) Or IsEmpty(numero_ddt.value)) Then
			sped.numero_ddt = ""
		Else 
			sped.numero_ddt = numero_ddt.value
		End If		
		
		If (sped.numero_ddt = "") Then
			Dim errclNDDT: Set errclNDDT = New errorclass
			errclNDDT.cod   = "000002"
			errclNDDT.tipo  = "REQUIRED"
			errclNDDT.field = "numero_ddt"
			errclNDDT.desc  = "IMPOSTARE NUMERO DDT"
			SpedErrorsAdd(errclNDDT)
			anyErrors = True  
		End If
				
		Set data_ddt = document.getElementById( "data_ddt" )
		
		If (IsNull(data_ddt.value) Or IsEmpty(data_ddt.value)) Then
			sped.data_ddt = ""
		Else 
			sped.data_ddt = data_ddt.value
		End If		
		
		If (sped.data_ddt = "") Then
			Dim errclDDDT: Set errclDDDT = New errorclass
			errclDDDT.cod   = "000003"
			errclDDDT.tipo  = "REQUIRED"
			errclDDDT.field = "data_ddt"
			errclDDDT.desc  = "IMPOSTARE DATA DDT"
			SpedErrorsAdd(errclDDDT)
			anyErrors = True  
		End If
		
		If (sped.data_ddt <> "") Then 
			Set objRE = New RegExp

			With objRE
				.Pattern    = "^(\d{1,2})/(\d{1,2})/(\d{4})$"
				.IgnoreCase = True
				.Global     = False
			End With
			
			If Not objRE.Test( sped.data_ddt ) Then
				Dim errclDDDTMF: Set errclDDDTMF = New errorclass
				errclDDDTMF.cod   = "000010"
				errclDDDTMF.tipo  = "MALFORMED"
				errclDDDTMF.field = "data_ddt"
				errclDDDTMF.desc  = "DATA DDT NON CORRETTA"
				SpedErrorsAdd(errclDDDTMF)
				anyErrors = True  
			End If
		End If 
			
		'Set titolo_oro_id = document.getElementById( "titolo_oro_id" )

		'If (IsNull(titolo_oro_id.value) Or IsEmpty(titolo_oro_id.value)) Then
		'	sped.titolo_oro_id = ""
		'Else 
		'	sped.titolo_oro_id  = titolo_oro_id.value
		'End If		

		'If (sped.titolo_oro_id = "") Then
		'	Dim errclTOI: Set errclTOI = New errorclass
		'	errclTOI.cod   = "000004"
		'	errclTOI.tipo  = "REQUIRED"
		'	errclTOI.field = "titolo_oro_id"
		'	errclTOI.desc  = "IMPOSTARE TITOLO ORO"
		'	SpedErrorsAdd(errclTOI)
		'	anyErrors = True  
		'End If
				
		Set totale_grammi_rottami = document.getElementById( "totale_grammi_rottami" )
		sped.totale_grammi_rottami = totale_grammi_rottami.value

		If (IsNull(totale_grammi_rottami.value) Or IsEmpty(totale_grammi_rottami.value)) Then
			sped.totale_grammi_rottami = 0
		Else 
			sped.totale_grammi_rottami  = CDbl(totale_grammi_rottami.value)  
		End If		

		If (sped.totale_grammi_rottami <= 0) Then
			Dim errclTGR: Set errclTGR = New errorclass
			errclTGR.cod   = "000005"
			errclTGR.tipo  = "GREATER_THEN"
			errclTGR.field = "totale_grammi_rottami"
			errclTGR.desc  = "IMPOSTARE TOTALE"
			SpedErrorsAdd(errclTGR)
			anyErrors = True  
		End If

		Set totale_grammi_puro_stimato = document.getElementById( "totale_grammi_puro_stimato" )

		If (IsNull(totale_grammi_puro_stimato.value) Or IsEmpty(totale_grammi_puro_stimato.value)) Then
			sped.totale_grammi_puro_stimato = 0
		Else 
			sped.totale_grammi_puro_stimato  = CDbl(totale_grammi_puro_stimato.value)  
		End If		

		'If (sped.totale_grammi_puro <= 0) Then
		'	Dim errclTGP: Set errclTGP = New errorclass
		'	errclTGP.cod   = "000006"
		'	errclTGP.tipo  = "GREATER_THEN"
		'	errclTGP.field = "totale_grammi_puro"
		'	errclTGP.desc  = "IMPOSTARE TOTALE"
		'	SpedErrorsAdd(errclTGP)
		'	anyErrors = True  
		'End If

		Set verga_stimata = document.getElementById( "verga_stimata" )
		sped.verga_stimata = verga_stimata.value
		
		Set titolo_stimato_verga = document.getElementById( "titolo_stimato_verga" )
		sped.titolo_stimato_verga = titolo_stimato_verga.value
		
		Set verga_fonderia = document.getElementById( "verga_fonderia" )
		
		If (IsNull(verga_fonderia.value) Or IsEmpty(verga_fonderia.value)) Then
			sped.verga_fonderia = 0
		Else 
			sped.verga_fonderia  = CDbl(verga_fonderia.value)  
		End If		

		If (sped.verga_fonderia <= 0) Then
			Dim errclVF: Set errclVF = New errorclass
			errclVF.cod   = "000007"
			errclVF.tipo  = "GREATER_THEN"
			errclVF.field = "verga_fonderia"
			errclVF.desc  = "IMPOSTARE VERGA"
			SpedErrorsAdd(errclVF)
			anyErrors = True  
		End If
				
		Set titolo_fonderia = document.getElementById( "titolo_fonderia" )
		sped.titolo_fonderia = titolo_fonderia.value

		If (IsNull(titolo_fonderia.value) Or IsEmpty(titolo_fonderia.value)) Then
			sped.titolo_fonderia = 0
		Else 
			sped.titolo_fonderia  = CDbl(titolo_fonderia.value)  
		End If		

		If (sped.titolo_fonderia <= 0) Then
			Dim errclTF: Set errclTF = New errorclass
			errclTF.cod   = "000008"
			errclTF.tipo  = "GREATER_THEN"
			errclTF.field = "titolo_fonderia"
			errclTF.desc  = "IMPOSTARE TITOLO"
			SpedErrorsAdd(errclTF)
			anyErrors = True  
		End If
		
		Set totale_grammi_puro_fonderia = document.getElementById( "totale_grammi_puro_fonderia" )
		sped.totale_grammi_puro_fonderia = totale_grammi_puro_fonderia.value
		
		Set titolo_lab_controsaggio = document.getElementById( "titolo_lab_controsaggio" )
		
		If (IsNull(titolo_lab_controsaggio.value) Or IsEmpty(titolo_lab_controsaggio.value)) Then
			sped.titolo_lab_controsaggio = 0
		Else 
			sped.titolo_lab_controsaggio  = CDbl(titolo_lab_controsaggio.value)  
		End If		

		If (sped.titolo_lab_controsaggio <= 0) Then
			Dim errclTLC: Set errclTLC = New errorclass
			errclTLC.cod   = "000009"
			errclTLC.tipo  = "GREATER_THEN"
			errclTLC.field = "titolo_lab_controsaggio"
			errclTLC.desc  = "IMPOSTARE TITOLO"
			SpedErrorsAdd(errclTLC)
			anyErrors = True  
		End If
				
		Set puro_stimato_lab_controsaggio = document.getElementById( "puro_stimato_lab_controsaggio" )
		sped.puro_stimato_lab_controsaggio = puro_stimato_lab_controsaggio.value
				
		Set NodeListDG = document.getElementsByName("differenza_grammi") 
 		For Each Elem In NodeListDG
		  	sped.differenza_grammi = Elem.innerHTML
 		Next
		Set NodeListDP = document.getElementsByName("differenza_percentuale") 
 		For Each Elem In NodeListDP
		  	sped.differenza_percentuale = Replace(Elem.innerHTML,"%","")
 		Next
		Set NodeListDGCS = document.getElementsByName("differenza_grammi_con_saggio") 
 		For Each Elem In NodeListDGCS
		  	sped.differenza_grammi_con_saggio = Elem.innerHTML
 		Next
		Set NodeListDPCS = document.getElementsByName("differenza_percentuale_con_saggio") 
 		For Each Elem In NodeListDPCS
		  	sped.differenza_percentuale_con_saggio = Replace(Elem.innerHTML,"%","")
 		Next

		If ( Not anyErrors And genKT) Then 
			If (sped.kt = "") Then
				sped.kt = CreateGUID()
				isNewRecord = True
			End If
		End If

    End Sub
		
	Sub SpedDetailDisplay(sped)
		'MsgBox "init SpedDetailDisplay"

		Set banco_metalli = document.getElementById( "banco_metalli" )
		For Each opt In banco_metalli.Options
  			If opt.Value = sped.banco_metalli_id Then
    			opt.Selected = True
  			Else
    			opt.Selected = False
  			End If
		Next		
		Set banco_metalli_error_list = document.getElementsByName("banco_metalli_error") 
 		For Each Elem In banco_metalli_error_list
 			Dim banco_metalli_error_object: Set banco_metalli_error_object = New errorclass			
 			If spederrorsdict.Exists("banco_metalli_id") Then
				Set banco_metalli_error_object = spederrorsdict.Item("banco_metalli_id")
		  		Elem.innerHTML = banco_metalli_error_object.desc
		  	Else 
		  		Elem.innerHTML = ""
			End if 
 		Next
		
		Set numero_ddt = document.getElementById( "numero_ddt" )
		numero_ddt.value = sped.numero_ddt

		Set numero_ddt_error_list = document.getElementsByName("numero_ddt_error") 
 		For Each Elem In numero_ddt_error_list
 			Dim numero_ddt_error_object: Set numero_ddt_error_object = New errorclass			
 			If spederrorsdict.Exists("numero_ddt") Then
				Set numero_ddt_error_object = spederrorsdict.Item("numero_ddt")
		  		Elem.innerHTML = numero_ddt_error_object.desc
		  	Else 
		  		Elem.innerHTML = ""
			End if 
 		Next

		Set data_ddt = document.getElementById( "data_ddt" )
		data_ddt.value = sped.data_ddt
		Set data_ddt_error_list = document.getElementsByName("data_ddt_error") 
 		For Each Elem In data_ddt_error_list
 			Dim data_ddt_error_object: Set data_ddt_error_object = New errorclass			
 			If spederrorsdict.Exists("data_ddt") Then
				Set data_ddt_error_object = spederrorsdict.Item("data_ddt")
		  		Elem.innerHTML = data_ddt_error_object.desc
		  	Else 
		  		Elem.innerHTML = ""
			End if 
 		Next

		'Set titolo_oro_id = document.getElementById( "titolo_oro_id" )
		'For Each opt In titolo_oro_id.Options
  		'	If opt.Value = sped.titolo_oro_id Then
    	'		opt.Selected = True
  		'	Else
    	'		opt.Selected = False
  		'	End If
		'Next
		'Set titolo_oro_id_error_list = document.getElementsByName("titolo_oro_id_error") 
 		'For Each Elem In titolo_oro_id_error_list
 		'	Dim titolo_oro_id_error_object: Set titolo_oro_id_error_object = New errorclass			
 		'	If spederrorsdict.Exists("titolo_oro_id") Then
		'		Set titolo_oro_id_error_object = spederrorsdict.Item("titolo_oro_id")
		'  		Elem.innerHTML = titolo_oro_id_error_object.desc
		'  	Else 
		'  		Elem.innerHTML = ""
		'	End if 
 		'Next

		Set verga_stimata = document.getElementById( "verga_stimata" )
		verga_stimata.value = CStr(sped.verga_stimata)
		Set titolo_stimato_verga = document.getElementById( "titolo_stimato_verga" )
		titolo_stimato_verga.value = CStr(sped.titolo_stimato_verga)
		
		Set totale_grammi_rottami = document.getElementById( "totale_grammi_rottami" )
		totale_grammi_rottami.value = CStr(sped.totale_grammi_rottami)

		Set totale_grammi_rottami_error_list = document.getElementsByName("totale_grammi_rottami_error") 
 		For Each Elem In totale_grammi_rottami_error_list
 			Dim totale_grammi_rottami_error_object: Set totale_grammi_rottami_error_object = New errorclass			
 			If spederrorsdict.Exists("totale_grammi_rottami") Then
				Set totale_grammi_rottami_error_object = spederrorsdict.Item("totale_grammi_rottami")
		  		Elem.innerHTML = totale_grammi_rottami_error_object.desc
		  	Else 
		  		Elem.innerHTML = ""
			End if 
 		Next
				
		Set totale_grammi_puro_stimato = document.getElementById( "totale_grammi_puro_stimato" )
		totale_grammi_puro_stimato.value = CStr(sped.totale_grammi_puro_stimato)
		
		'Set totale_grammi_puro_error_list = document.getElementsByName("totale_grammi_puro_error") 
 		'For Each Elem In totale_grammi_puro_error_list
 		'	Dim totale_grammi_puro_error_object: Set totale_grammi_puro_error_object = New errorclass			
 		'	If spederrorsdict.Exists("totale_grammi_puro") Then
		'		Set totale_grammi_puro_error_object = spederrorsdict.Item("totale_grammi_puro")
		' 		Elem.innerHTML = totale_grammi_puro_error_object.desc
		'  	Else 
		' 		Elem.innerHTML = ""
		'	End if 
 		'Next

		Set verga_fonderia = document.getElementById( "verga_fonderia" )
		verga_fonderia.value = CStr(sped.verga_fonderia)
		
		Set verga_fonderia_error_list = document.getElementsByName("verga_fonderia_error") 
 		For Each Elem In verga_fonderia_error_list
 			Dim verga_fonderia_error_object: Set verga_fonderia_error_object = New errorclass			
 			If spederrorsdict.Exists("verga_fonderia") Then
				Set verga_fonderia_error_object = spederrorsdict.Item("verga_fonderia")
		  		Elem.innerHTML = verga_fonderia_error_object.desc
		  	Else 
		  		Elem.innerHTML = ""
			End if 
 		Next
		
		Set titolo_fonderia = document.getElementById( "titolo_fonderia" )
		titolo_fonderia.value = CStr(sped.titolo_fonderia)
		
		Set titolo_fonderia_error_list = document.getElementsByName("titolo_fonderia_error") 
 		For Each Elem In titolo_fonderia_error_list
 			Dim titolo_fonderia_error_object: Set titolo_fonderia_error_object = New errorclass			
 			If spederrorsdict.Exists("titolo_fonderia") Then
				Set titolo_fonderia_error_object = spederrorsdict.Item("titolo_fonderia")
		  		Elem.innerHTML = titolo_fonderia_error_object.desc
		  	Else 
		  		Elem.innerHTML = ""
			End if 
 		Next
		
		Set totale_grammi_puro_fonderia = document.getElementById( "totale_grammi_puro_fonderia" )
		totale_grammi_puro_fonderia.value = CStr(sped.totale_grammi_puro_fonderia) 
		
		Set titolo_lab_controsaggio = document.getElementById( "titolo_lab_controsaggio" )
		titolo_lab_controsaggio.value = CStr(sped.titolo_lab_controsaggio)
		
		Set titolo_lab_controsaggio_error_list = document.getElementsByName("titolo_lab_controsaggio_error") 
 		For Each Elem In titolo_lab_controsaggio_error_list
 			Dim titolo_lab_controsaggio_error_object: Set titolo_lab_controsaggio_error_object = New errorclass			
 			If spederrorsdict.Exists("titolo_lab_controsaggio") Then
				Set titolo_lab_controsaggio_error_object = spederrorsdict.Item("titolo_lab_controsaggio")
		  		Elem.innerHTML = titolo_lab_controsaggio_error_object.desc
		  	Else 
		  		Elem.innerHTML = ""
			End if 
 		Next		
		
		Set puro_stimato_lab_controsaggio = document.getElementById( "puro_stimato_lab_controsaggio" )
		puro_stimato_lab_controsaggio.value = CStr(sped.puro_stimato_lab_controsaggio)
		
		'Set differenza_percentuale = document.getElementById( "differenza_percentuale" )
		'Set differenza_grammi_con_saggio = document.getElementById( "differenza_grammi_con_saggio" )
		'Set differenza_percentuale_con_saggio = document.getElementById( "differenza_percentuale_con_saggio" )
		
		Set NodeListDG = document.getElementsByName("differenza_grammi") 
 		For Each Elem In NodeListDG
		  	Elem.innerHTML =  CStr(sped.differenza_grammi)
 		Next
		Set NodeListDP = document.getElementsByName("differenza_percentuale") 
 		For Each Elem In NodeListDP
		  	Elem.innerHTML =  CStr(sped.differenza_percentuale) + "%"
 		Next
		Set NodeListDGCS = document.getElementsByName("differenza_grammi_con_saggio") 
 		For Each Elem In NodeListDGCS
		  	Elem.innerHTML =  CStr(sped.differenza_grammi_con_saggio)
 		Next
		Set NodeListDPCS = document.getElementsByName("differenza_percentuale_con_saggio") 
 		For Each Elem In NodeListDPCS
		  	Elem.innerHTML =  CStr(sped.differenza_percentuale_con_saggio) + "%"
 		Next
		
		Set kt = document.getElementById( "kt" )
		kt.value = sped.kt
		
		displayAllTitoli()
		
    End Sub
	
	Sub getStoreBM()
		Dim bm: Set bm = New bmclass
		With bm
				.kt = "c07cb0bc-f3c7-461f-ae7f-93b1430912db"
    			.desc = "Fonderia Pinco"
			End With
			bmdict.Add bm.kt, bm

			Set bm = new bmclass
			With bm
				.kt = "a799419d-ae66-4174-986e-1da78274695a"
    			.desc = "Fonderia Banco Metalli Vicenza"
			End With
			bmdict.Add bm.kt, bm

    End Sub

	Sub getStoreTO()
		Dim toi: Set toi = new toclass
		With toi
				.kt = "x07cb0bc-f3c7-461f-ae7f-93b1430912db"
    			.desc = "999,9"
    			.coefficiente = 0.999
    			.coefficiente_titolo_stimato = 0.999
			End With
			todict.Add toi.kt, toi

			Set toi = new toclass
			With toi
				.kt = "y799419d-ae66-4174-986e-1da78274695a"
    			.desc = "985"
    			.coefficiente = 0.98
    			.coefficiente_titolo_stimato = 0.998
			End With
			todict.Add toi.kt, toi

			Set toi = new toclass
			With toi
				.kt = "z71d8c09-2351-4c04-8087-c9d7c7876e12"
    			.desc = "916"
    			.coefficiente = 0.914
    			.coefficiente_titolo_stimato = 0.998
			End With
			todict.Add toi.kt, toi

			Set toi = new toclass
			With toi
				.kt = "y71d8c09-2351-4c04-8087-c9d7c7876e12"
    			.desc = "900"
    			.coefficiente = 0.894    			
    			.coefficiente_titolo_stimato = 0.998
			End With
			todict.Add toi.kt, toi

			Set toi = new toclass
			With toi
				.kt = "w71d8c09-2351-4c04-8087-c9d7c7876e12"
    			.desc = "750"
    			.coefficiente = 0.738  			
    			.coefficiente_titolo_stimato = 0.992
			End With
			todict.Add toi.kt, toi

			Set toi = new toclass
			With toi
				.kt = "a71d8c09-2351-4c04-8087-c9d7c7876e12"				
    			.desc = "585"
    			.coefficiente = 0.55    			
    			.coefficiente_titolo_stimato = 0.99
			End With
			todict.Add toi.kt, toi

			Set toi = new toclass
			With toi
				.kt = "b71d8c09-2351-4c04-8087-c9d7c7876e12"
    			.desc = "500"
    			.coefficiente = 0.475
    			.coefficiente_titolo_stimato = 0.99    			
			End With
			todict.Add toi.kt, toi

			Set toi = new toclass
			With toi
				.kt = "c71d8c09-2351-4c04-8087-c9d7c7876e12"
    			.desc = "375"
    			.coefficiente = 0.35
    			.coefficiente_titolo_stimato = 0.985
			End With
			todict.Add toi.kt, toi

			Set toi = new toclass
			With toi
				.kt = "d71d8c09-2351-4c04-8087-c9d7c7876e12"
    			.desc = "333"
    			.coefficiente = 0.318    			
    			.coefficiente_titolo_stimato = 0.985
			End With
			todict.Add toi.kt, toi

    End Sub

	Sub getStoredSpedDetts()	
		Dim speddet: Set speddet = new speddetclass

		'titolo oro desc => 750
		With speddet
			.kt = "94da5c56-0c30-477a-bc24-ac603a30e3c7"
			.fk = "edfc686e267e4a8daa434ee9577e81c8"
    		.titolo_oro_id = "w71d8c09-2351-4c04-8087-c9d7c7876e12"
    		.grammi_lordi = 3015.44
		End With	
		speddetdict.Add speddet.kt, speddet
		
		Set speddet = new speddetclass
		'titolo oro desc => 750
		With speddet
			.kt = "cd1cdfa0-bafb-4362-b3ca-0efa035bee97"
			.fk = "bf7431f2b63449569d067c5705d24a67"
    		.titolo_oro_id = "w71d8c09-2351-4c04-8087-c9d7c7876e12"
    		.grammi_lordi = 3030.21
		End With	
		speddetdict.Add speddet.kt, speddet
		
		Set speddet = new speddetclass
		'titolo oro desc => 750
		With speddet
			.kt = "32cf79e5-a2c8-4e22-9273-6fd04cde5a2c"
			.fk = "a067b56f757841cb93af2a0482eb4451"
    		.titolo_oro_id = "w71d8c09-2351-4c04-8087-c9d7c7876e12"
    		.grammi_lordi = 2988.5
		End With	
		speddetdict.Add speddet.kt, speddet
    End Sub

	Sub getDettsOfSped(fk)
		currspeddetdict.RemoveAll()
		If ( fk <> "" ) Then 
			For Each i In speddetdict.Keys
    			Set csddi = speddetdict.Item(i)
    			If ( csddi.fk = fk ) Then
    				currspeddetdict.Add csddi.kt, csddi
    			End If 
			Next
		End If      	
	End Sub 
	
	Sub getStoredSpeds()
			Dim sped: Set sped = new spedclass
			Dim bm: Set bm = new bmclass
			Dim toi: Set toi = new toclass
			With sped
				.kt = "edfc686e267e4a8daa434ee9577e81c8"
    			.banco_metalli_id = "c07cb0bc-f3c7-461f-ae7f-93b1430912db"
    			'.titolo_oro_id = "w71d8c09-2351-4c04-8087-c9d7c7876e12"
    			.data_ddt = "17/11/2022"
    			.numero_ddt = 122
				.totale_grammi_puro_stimato = 2225.39
				.verga_stimata = 2994.33
				.titolo_stimato_verga = 743   
    			.totale_grammi_rottami = 3015.44
    			.verga_fonderia = 2973.1
    			.titolo_fonderia = 739
    			.titolo_lab_controsaggio = 742
			End With

			speddict.Add sped.kt, sped
			Set sped = new spedclass
			With sped
				.kt = "bf7431f2b63449569d067c5705d24a67"
    			.banco_metalli_id = "a799419d-ae66-4174-986e-1da78274695a"
    			'.titolo_oro_id = "w71d8c09-2351-4c04-8087-c9d7c7876e12"
    			.data_ddt = "24/11/2022"
    			.numero_ddt = 123
				.totale_grammi_puro_stimato = 2236.29498
				.verga_stimata = 3009
				.titolo_stimato_verga = 743   
    			.totale_grammi_rottami = 3030.21
    			.verga_fonderia = 3008.9
    			.titolo_fonderia = 742
    			.titolo_lab_controsaggio = 743
			End With

			speddict.Add sped.kt, sped
			Set sped = new spedclass
			With sped
				.kt = "a067b56f757841cb93af2a0482eb4451"
    			.banco_metalli_id = "c07cb0bc-f3c7-461f-ae7f-93b1430912db"
    			'.titolo_oro_id = "w71d8c09-2351-4c04-8087-c9d7c7876e12"
    			.data_ddt = "02/12/2022"
    			.numero_ddt = 124
				.totale_grammi_puro_stimato = 2205.51
				.verga_stimata = 2967.58
				.titolo_stimato_verga = 743   
    			.totale_grammi_rottami = 2988.5
    			.verga_fonderia = 2965.9
    			.titolo_fonderia = 740
    			.titolo_lab_controsaggio = 743    			
			End With

			speddict.Add sped.kt, sped
    End Sub


	Sub BuildSelectTO(node)
		Dim options: Set options = node.getElementsByTagName("option")
		For Each option_item In options
    		node.removeChild(option_item)
		Next

    	Set optNodeVoid = document.createElement("option")
    	Set attrVoid = document.createAttribute("value")
		attrVoid.value = ""
		optNodeVoid.setAttributeNode(attrVoid)
		optNodeVoid.innerHTML = ""
    	node.appendChild(optNodeVoid)

		For Each i In todict.Keys
    		Set toi = todict.Item(i)

    		Set optNode = document.createElement("option")
    		Set attr = document.createAttribute("value")
			attr.value = toi.kt
			optNode.setAttributeNode(attr)
			optNode.innerHTML = toi.desc
    		node.appendChild(optNode)
		Next		
    End Sub

	Sub BuildSelectBM(node)
		Dim options: Set options = node.getElementsByTagName("option")
		For Each option_item In options
    		node.removeChild(option_item)
		Next

    	Set optNodeVoid = document.createElement("option")
    	Set attrVoid = document.createAttribute("value")
		attrVoid.value = ""
		optNodeVoid.setAttributeNode(attrVoid)
		optNodeVoid.innerHTML = ""
    	node.appendChild(optNodeVoid)

		For Each i In bmdict.Keys
    		Set bm = bmdict.Item(i)

    		Set optNode = document.createElement("option")
    		Set attr = document.createAttribute("value")
			attr.value = bm.kt
			optNode.setAttributeNode(attr)
			optNode.innerHTML = bm.desc
    		node.appendChild(optNode)
		Next		
    End Sub
	
	Sub displaySped(sped)
	    	Set tableNode = document.getElementById( "spedizioni_table" )
    		Set trNode = document.createElement("tr")
    		Set attr = document.createAttribute("class")
			attr.value = "spedrow"
			trNode.setAttributeNode(attr)

    		Set attrClassField = document.createAttribute("class")
			attrClassField.value = "spedizioni_field"
    		
    		Set tdNodeBM = document.createElement("td")
    		tdNodeBM.innerHTML = "<p> " + CStr(sped.banco_metalli_desc) + " </p> "    		
    		Set attrClassFieldBM = document.createAttribute("class")
			attrClassFieldBM.value = "spedizioni_field"
			tdNodeBM.setAttributeNode(attrClassFieldBM)
    		trNode.appendChild(tdNodeBM)
    		
    		'Set tdNodeTO = document.createElement("td")
    		'tdNodeTO.innerHTML = "<p> " + CStr(sped.titolo_oro_desc) + " </p> "
    		'Set attrClassFieldTO = document.createAttribute("class")
			'attrClassFieldTO.value = "spedizioni_field"
			'tdNodeTO.setAttributeNode(attrClassFieldTO)
    		'trNode.appendChild(tdNodeTO)

    		Set tdNodeDDT = document.createElement("td")
    		tdNodeDDT.innerHTML = "<p> " + CStr(sped.numero_ddt) + " </p> "
    		Set attrClassFieldDDT = document.createAttribute("class")
			attrClassFieldDDT.value = "spedizioni_field"
			tdNodeDDT.setAttributeNode(attrClassFieldDDT)
    		trNode.appendChild(tdNodeDDT)

    		Set tdNodeDataDDT = document.createElement("td")
    		tdNodeDataDDT.innerHTML = "<p> " + CStr(sped.data_ddt) + " </p> "
    		Set attrClassFieldDataDDT = document.createAttribute("class")
			attrClassFieldDataDDT.value = "spedizioni_field"
			tdNodeDataDDT.setAttributeNode(attrClassFieldDataDDT)
    		trNode.appendChild(tdNodeDataDDT)
    		    		
    		'Set tdNodeGL = document.createElement("td")
    		'tdNodeGL.innerHTML = "<p> " + CStr(sped.grammi_lordi) + " </p> "
    		'Set attrClassFieldGL = document.createAttribute("class")
			'attrClassFieldGL.value = "spedizioni_field"
			'tdNodeGL.setAttributeNode(attrClassFieldGL)
    		'trNode.appendChild(tdNodeGL)

    		'Set tdNodeGPS = document.createElement("td")
    		'tdNodeGPS.innerHTML = "<p> " + CStr(sped.grammi_puri_stimati) + " </p> "
    		'Set attrClassFieldGPS = document.createAttribute("class")
			'attrClassFieldGPS.value = "spedizioni_field"
			'tdNodeGPS.setAttributeNode(attrClassFieldGPS)
    		'trNode.appendChild(tdNodeGPS)
    		
    		Set tdNodeVS = document.createElement("td")
    		tdNodeVS.innerHTML = "<p> " + CStr(sped.verga_stimata) + " </p> "
    		Set attrClassFieldVS = document.createAttribute("class")
			attrClassFieldVS.value = "spedizioni_field"
			tdNodeVS.setAttributeNode(attrClassFieldVS)
    		trNode.appendChild(tdNodeVS)
    		
    		Set tdNodeTSV = document.createElement("td")
    		tdNodeTSV.innerHTML = "<p> " + CStr(sped.titolo_stimato_verga) + " </p> "
    		Set attrClassFieldTSV = document.createAttribute("class")
			attrClassFieldTSV.value = "spedizioni_field"
			tdNodeTSV.setAttributeNode(attrClassFieldTSV)
    		trNode.appendChild(tdNodeTSV)
    		
    		Set tdNodeTGR = document.createElement("td")
    		tdNodeTGR.innerHTML = "<p> " + CStr(sped.totale_grammi_rottami) + " </p> "
    		Set attrClassFieldTGR = document.createAttribute("class")
			attrClassFieldTGR.value = "spedizioni_field"
			tdNodeTGR.setAttributeNode(attrClassFieldTGR)
    		trNode.appendChild(tdNodeTGR)

    		Set tdNodeTGPS = document.createElement("td")
    		tdNodeTGPS.innerHTML = "<p> " + CStr(sped.totale_grammi_puro_stimato) + " </p> "
    		Set attrClassFieldTGPS = document.createAttribute("class")
			attrClassFieldTGPS.value = "spedizioni_field"
			tdNodeTGPS.setAttributeNode(attrClassFieldTGPS)
    		trNode.appendChild(tdNodeTGPS)

    		Set tdNodeVIP = document.createElement("td")
    		tdNodeVIP.innerHTML = "<p> " + CStr(sped.verga_fonderia) + " </p> "
    		Set attrClassFieldVIP = document.createAttribute("class")
			attrClassFieldVIP.value = "spedizioni_field"
			tdNodeVIP.setAttributeNode(attrClassFieldVIP)
    		trNode.appendChild(tdNodeVIP)

    		Set tdNodeTIP = document.createElement("td")
    		tdNodeTIP.innerHTML = "<p> " + CStr(sped.titolo_fonderia) + " </p> "
    		Set attrClassFieldTIP = document.createAttribute("class")
			attrClassFieldTIP.value = "spedizioni_field"
			tdNodeTIP.setAttributeNode(attrClassFieldTip)
    		trNode.appendChild(tdNodeTIP)
    		
    		Set tdNodeTGPIP = document.createElement("td")
    		tdNodeTGPIP.innerHTML = " <p> " + CStr(sped.totale_grammi_puro_fonderia) + " </p> "
    		Set attrClassFieldTGPIP = document.createAttribute("class")
			attrClassFieldTGPIP.value = "spedizioni_field"
			tdNodeTGPIP.setAttributeNode(attrClassFieldTGPIP)
    		trNode.appendChild(tdNodeTGPIP)

    		Set tdNodeTLC = document.createElement("td")
    		tdNodeTLC.innerHTML = "<p> " + CStr(sped.titolo_lab_controsaggio) + " </p> "
    		Set attrClassFieldTLC = document.createAttribute("class")
			attrClassFieldTLC.value = "spedizioni_field"
			tdNodeTLC.setAttributeNode(attrClassFieldTLC)
    		trNode.appendChild(tdNodeTLC)
    		
    		Set tdNodeTGPLC = document.createElement("td")
    		tdNodeTGPLC.innerHTML = " <p> " + CStr(sped.puro_stimato_lab_controsaggio) + " </p> "
    		Set attrClassFieldTGPLC = document.createAttribute("class")
			attrClassFieldTGPLC.value = "spedizioni_field"
			tdNodeTGPLC.setAttributeNode(attrClassFieldTGPLC)
    		trNode.appendChild(tdNodeTGPLC)
    		
    		Set tdNodeDiffGR = document.createElement("td")
    		tdNodeDiffGR.innerHTML = sped.differenza_grammi
    		Set attrStyleGR = document.createAttribute("style")
			attrStyleGR.value = "background-color:#fb923c;color:white;font-weigth:bolder;text-align:center;"
			tdNodeDiffGR.setAttributeNode(attrStyleGR)
    		trNode.appendChild(tdNodeDiffGR)    		

    		Set tdNodeDiffPERC = document.createElement("td")
    		tdNodeDiffPERC.innerHTML = CStr(sped.differenza_percentuale) + "%"
    		Set attrStylePERC = document.createAttribute("style")
			attrStylePERC.value = "background-color:#fb923c;color:white;font-weigth:bolder;text-align:center;"
			tdNodeDiffPERC.setAttributeNode(attrStylePERC)
    		trNode.appendChild(tdNodeDiffPERC)    		

    		Set tdNodeDiffGRCS = document.createElement("td")
    		tdNodeDiffGRCS.innerHTML = sped.differenza_grammi_con_saggio
    		Set attrStyleGRCS = document.createAttribute("style")
			attrStyleGRCS.value = "background-color:#fb923c;color:white;font-weigth:bolder;text-align:center;"
			tdNodeDiffGRCS.setAttributeNode(attrStyleGRCS)
    		trNode.appendChild(tdNodeDiffGRCS)    		

    		Set tdNodeDiffPERCCS = document.createElement("td")
    		tdNodeDiffPERCCS.innerHTML = CStr(sped.differenza_percentuale_con_saggio) + "%"
    		Set attrStylePERCCS = document.createAttribute("style")
			attrStylePERCCS.value = "background-color:#fb923c;color:white;font-weigth:bolder;text-align:center;"
			tdNodeDiffPERCCS.setAttributeNode(attrStylePERCCS)
    		trNode.appendChild(tdNodeDiffPERCCS)    		

    		Set iRNode = document.createElement("i")
    		Set attrClass = document.createAttribute("class")
			attrClass.value = "fa fa-sharp fa-solid fa-trash icon_style"
			iRNode.setAttributeNode(attrClass)
    		Set attrOnClick = document.createAttribute("onClick")
			attrOnClick.value = "delete_sped_js('" + CStr(sped.kt) + "')"
			iRNode.setAttributeNode(attrOnClick)			
    		Set tdNodeIR = document.createElement("td")
    		tdNodeIR.appendChild(iRNode) 
    		trNode.appendChild(tdNodeIR) 

    		Set iMNode = document.createElement("i")
    		Set attrClass = document.createAttribute("class")
			attrClass.value = "fa fa-sharp fa-solid fa-pencil icon_style"
			iMNode.setAttributeNode(attrClass)
    		Set attrOnClick = document.createAttribute("onClick")
			attrOnClick.value = "modify_sped('" + CStr(sped.kt) + "')"
			iMNode.setAttributeNode(attrOnClick)			
    		Set tdNodeMR = document.createElement("td")
    		tdNodeMR.appendChild(iMNode) 
    		trNode.appendChild(tdNodeMR) 
 
    		tableNode.appendChild(trNode)
	End Sub

	Sub calcSped(sped) 
		If bmdict.Exists(sped.banco_metalli_id) Then
			Set bm = bmdict.Item(sped.banco_metalli_id)
			sped.banco_metalli_desc = bm.desc
		Else 
			sped.banco_metalli_desc = ""
		End If
	
		If (  sped.verga_fonderia <> 0 And sped.titolo_fonderia <> 0 ) Then
    		sped.totale_grammi_puro_fonderia = Round(( sped.verga_fonderia * sped.titolo_fonderia ) / 1000,2)
    	Else
    		sped.totale_grammi_puro_fonderia = 0 
    	End If
    	sped.differenza_grammi = Round((sped.totale_grammi_puro_fonderia - sped.totale_grammi_puro_stimato),2)
		If (  sped.differenza_grammi <> 0 And sped.totale_grammi_puro_stimato <> 0 ) Then
        	sped.differenza_percentuale = Round((sped.differenza_grammi / sped.totale_grammi_puro_stimato) * 100 ,2)
        Else 
        	sped.differenza_percentuale = 0
        End If
		If (  sped.verga_fonderia <> 0 And sped.titolo_lab_controsaggio <> 0 ) Then
	    	sped.puro_stimato_lab_controsaggio = Round(( sped.verga_fonderia * sped.titolo_lab_controsaggio ) / 1000,4)
	    Else 
	    	sped.puro_stimato_lab_controsaggio = 0
	    End If
    	sped.differenza_grammi_con_saggio = Round(sped.totale_grammi_puro_fonderia - sped.puro_stimato_lab_controsaggio,2)
    		
		If (  sped.differenza_grammi_con_saggio <> 0 And sped.totale_grammi_puro_stimato <> 0 ) Then  
    		sped.differenza_percentuale_con_saggio = Round((sped.differenza_grammi_con_saggio / sped.totale_grammi_puro_stimato) * 100 ,2)
    	Else 
    		sped.differenza_percentuale_con_saggio = 0
    	End If
	End Sub

	Sub displayAllTitoli()
		Dim dett_spedizioni_table: Set dett_spedizioni_table = document.getElementById( "dett_spedizioni_table" )
		clean_table(dett_spedizioni_table)
		For Each i In todict.Keys
    		Set titolo = todict.Item(i)
    		displayTitolo(titolo)
		Next
    End Sub

	Sub displayTitolo(titolo)
	    Set tableNode = document.getElementById( "dett_spedizioni_table" )
    	Set trNode = document.createElement("tr")
    	Set attr = document.createAttribute("class")
		attr.value = "speddettrow"
		trNode.setAttributeNode(attr)

    	Set attrClassField = document.createAttribute("class")
		attrClassField.value = "spedizioni_field"
    		
    	Set tdNodeT = document.createElement("td")
    	tdNodeT.innerHTML = "<p> " + titolo.desc + " </p> "    		
    	Set attrClassFieldT = document.createAttribute("class")
		attrClassFieldT.value = "dett_spedizioni_field"
		tdNodeT.setAttributeNode(attrClassFieldT)
    	trNode.appendChild(tdNodeT)

    	Set tdNodeGL = document.createElement("td")
    	Set inputGL = document.createElement("input")
		inputGL.setAttribute "type", "text"
		inputGL.value = 0
    	tdNodeGL.appendChild(inputGL)    		
    	Set attrClassFieldGL = document.createAttribute("class")
		attrClassFieldGL.value = "dett_spedizioni_field"
		tdNodeGL.setAttributeNode(attrClassFieldGL)
    	trNode.appendChild(tdNodeGL)
    		
    	Set tdNodeGPS = document.createElement("td")
    	Set inputGPS = document.createElement("input")
		inputGPS.setAttribute "type", "text"
		inputGPS.setAttribute "disabled", ""
		inputGPS.value = 0
    	tdNodeGPS.appendChild(inputGPS)    		
    	Set attrClassFieldGPS = document.createAttribute("class")
		attrClassFieldGPS.value = "dett_spedizioni_field"
		tdNodeT.setAttributeNode(attrClassFieldGPS)
    	trNode.appendChild(tdNodeGPS)
    		
    	tableNode.appendChild(trNode)
	End Sub

	Sub displayAllSpeds()
		For Each i In speddict.Keys
    		Set sped = speddict.Item(i)
    		calcSped(sped)
    		displaySped(sped)
		Next
    End Sub
	
	Sub show_speds_list() 
    	document.getElementById("list_speds_container").style.display="block"
    	document.getElementById("mod_add_sped_container").style.display="none"
    End Sub
	
	Sub show_mod_ins_sped() 
    	document.getElementById("list_speds_container").style.display="none"
    	document.getElementById("mod_add_sped_container").style.display="block"
    End Sub

	Sub insert_sped()
		Dim sped: Set sped = new spedclass		
		SpedErrorsCleared()
		
		sped.kt = ""
		sped.banco_metalli_id = ""
		'sped.titolo_oro_id = ""
		sped.data_ddt = ""
		sped.numero_ddt  = ""
		'sped.grammi_lordi = 0
		'sped.grammi_puri_stimati = 0
		sped.totale_grammi_rottami = 0
		sped.totale_grammi_puro_stimato = 0
		sped.verga_stimata = 0
    	sped.titolo_stimato_verga = 0
    	sped.totale_grammi_puro_fonderia = 0
    	sped.differenza_grammi = 0
    	sped.differenza_percentuale = 0
    	sped.differenza_grammi_con_saggio = 0
    	sped.differenza_percentuale_con_saggio = 0
    	sped.verga_fonderia = 0
    	sped.titolo_fonderia = 0
    	sped.titolo_lab_controsaggio = 0 
    	sped.puro_stimato_lab_controsaggio = 0
    	
    	getDettsOfSped("")
    	
		add_sped()
		SpedDetailDisplay(sped)
	End Sub
	
	Sub add_sped()
    	'Dim tableNode, trNode , tdNode
  		'Dim TheForm
  		'Set TheForm = Document.forms("ValidForm")
    	'Dim numero_ddt: Set numero_ddt = TheForm.elements.numero_ddt
    	'Dim inputs: Set inputs = TheForm.getElementsByTagName("input")
    	
    	show_mod_ins_sped()
    	Dim banco_metalli: Set banco_metalli = document.getElementById("banco_metalli")
    	BuildSelectBM(banco_metalli)
    	'Dim titolo_oro: Set titolo_oro = document.getElementById("titolo_oro_id")
    	'BuildSelectTO(titolo_oro)
    	
    End Sub
    
	Function CreateGUID
  		Dim TypeLib
  		Set TypeLib = CreateObject("Scriptlet.TypeLib")
  		CreateGUID = LCase(Mid(TypeLib.Guid, 2, 36))
	End Function
	
    Sub submit_sped
  		'Dim TheForm
  		'Set TheForm = Document.forms("ValidForm")
  		Dim existErrors
		existErrors = False 
		SpedErrorsCleared()
			
		Dim sped: Set sped = new spedclass
		Dim isNewRecord
		isNewRecord = False
				
		Call SpedDetailGet(sped,isNewRecord,True)
		calcSped(sped) 
		existErrors = SpedErrorsStatus()

		If (Not existErrors) Then		
			If (Not isNewRecord) Then 
				'MsgBox "modifico record sped con chiave '" + sped.kt + "'"
				If speddict.Exists(sped.kt) Then
					Set spedDest = speddict.Item(sped.kt)
					Call SpedDetailCopy(spedDest,sped)
				Else 
					MsgBox "Non esiste una spedizione con chiave tecnica '" + kt + "'"
				End If
			Else
				'MsgBox "aggiungo record sped con chiave '" + sped.kt + "'"
 				speddict.Add sped.kt, sped
			End If 

			Dim spedizioni_table: Set spedizioni_table = document.getElementById( "spedizioni_table" )
			clean_table(spedizioni_table)
			displayAllSpeds()
			show_speds_list()
		Else 
			SpedDetailDisplay(sped)
			MsgBox "ci sono errori"
		End If 
		
  		'If IsNumeric(TheForm.Text1.Value) Then
    	'	If TheForm.Text1.Value < 1 Or TheForm.Text1.Value > 10 Then
      	'		MsgBox "Please enter a number between 1 and 10."
    	'	Else
      	'		MsgBox "Thank you."
    	'	End If
  		'Else
    	'	MsgBox "Please enter a numeric value."
  		'End If
	End Sub

	Sub del_from_dict_sped(kt)
		If speddict.Exists(kt) Then
			speddict.Remove(kt)
		End if
	End Sub
	
	Sub clean_table(node)
		Dim trs: Set trs = node.getElementsByTagName("tr")
		For Each tr_item In trs
			Dim classItemAttr: Set classItemAttr = tr_item.attributes.getNamedItem("class")
			If classItemAttr.value <> "header_tr" Then 
    			node.removeChild(tr_item)
    		End if
		Next
	End Sub
		
	Sub delete_sped(kt)
		del_from_dict_sped(kt)
		Dim spedizioni_table: Set spedizioni_table = document.getElementById( "spedizioni_table" )
		clean_table(spedizioni_table)
		displayAllSpeds()
	End Sub

	Sub modify_sped(kt)
		If speddict.Exists(kt) Then
			SpedErrorsCleared()
			Set sped = speddict.Item(kt)
			getDettsOfSped(kt)
			add_sped()
			SpedDetailDisplay(sped)
		Else 
			MsgBox "Non esiste una spedizione con chiave tecnica '" + kt + "'"
		End If
	End Sub

	Sub select_tab_speds()
	
	    Dim navbar_item_sped: Set navbar_item_sped = document.getElementById( "navbar_item_sped" )
	    Dim spedizioni: Set spedizioni = document.getElementById( "spedizioni_section" )

	    Set attrClassSped = document.createAttribute("class")
		attrClassSped.value = "navbar_item navbar_item_selected"
		navbar_item_sped.setAttributeNode(attrClassSped)

	    Set attrStyleSped = document.createAttribute("style")
		attrStyleSped.value = "display:block;"
		spedizioni.setAttributeNode(attrStyleSped)

	    Dim navbar_item_bm: Set navbar_item_bm = document.getElementById( "navbar_item_bm" )
	    Dim banco_metalli: Set banco_metalli = document.getElementById( "banco_metalli_section" )

	    Set attrClassBM = document.createAttribute("class")
		attrClassBM.value = "navbar_item "
		navbar_item_bm.setAttributeNode(attrClassBM)

	    Set attrStyleBM = document.createAttribute("style")
		attrStyleBM.value = "display:none;"
		banco_metalli.setAttributeNode(attrStyleBM)
		
	    Dim navbar_item_titoli: Set navbar_item_titoli = document.getElementById( "navbar_item_titoli" )
	    Dim titoli: Set titoli = document.getElementById( "titoli_section" )

	    Set attrClassTitoli = document.createAttribute("class")
		attrClassTitoli.value = "navbar_item "
		navbar_item_titoli.setAttributeNode(attrClassTitoli)

	    Set attrStyleTitoli = document.createAttribute("style")
		attrStyleTitoli.value = "display:none;"
		titoli.setAttributeNode(attrStyleTitoli)

	End Sub

	Sub select_tab_bm()
	    Dim navbar_item_sped: Set navbar_item_sped = document.getElementById( "navbar_item_sped" )
	    Dim spedizioni: Set spedizioni = document.getElementById( "spedizioni_section" )

	    Set attrClassSped = document.createAttribute("class")
		attrClassSped.value = "navbar_item"
		navbar_item_sped.setAttributeNode(attrClassSped)

	    Set attrStyleSped = document.createAttribute("style")
		attrStyleSped.value = "display:none;"
		spedizioni.setAttributeNode(attrStyleSped)

	    Dim navbar_item_bm: Set navbar_item_bm = document.getElementById( "navbar_item_bm" )
	    Dim banco_metalli: Set banco_metalli = document.getElementById( "banco_metalli_section" )

	    Set attrClassBM = document.createAttribute("class")
		attrClassBM.value = "navbar_item navbar_item_selected"
		navbar_item_bm.setAttributeNode(attrClassBM)

	    Set attrStyleBM = document.createAttribute("style")
		attrStyleBM.value = "display:block;"
		banco_metalli.setAttributeNode(attrStyleBM)
		
	    Dim navbar_item_titoli: Set navbar_item_titoli = document.getElementById( "navbar_item_titoli" )
	    Dim titoli: Set titoli = document.getElementById( "titoli_section" )

	    Set attrClassTitoli = document.createAttribute("class")
		attrClassTitoli.value = "navbar_item "
		navbar_item_titoli.setAttributeNode(attrClassTitoli)

	    Set attrStyleTitoli = document.createAttribute("style")
		attrStyleTitoli.value = "display:none;"
		titoli.setAttributeNode(attrStyleTitoli)
	End Sub

	Sub select_tab_titoli()
	    Dim navbar_item_sped: Set navbar_item_sped = document.getElementById( "navbar_item_sped" )
	    Dim spedizioni: Set spedizioni = document.getElementById( "spedizioni_section" )

	    Set attrClassSped = document.createAttribute("class")
		attrClassSped.value = "navbar_item"
		navbar_item_sped.setAttributeNode(attrClassSped)

	    Set attrStyleSped = document.createAttribute("style")
		attrStyleSped.value = "display:none;"
		spedizioni.setAttributeNode(attrStyleSped)

	    Dim navbar_item_bm: Set navbar_item_bm = document.getElementById( "navbar_item_bm" )
	    Dim banco_metalli: Set banco_metalli = document.getElementById( "banco_metalli_section" )

	    Set attrClassBM = document.createAttribute("class")
		attrClassBM.value = "navbar_item"
		navbar_item_bm.setAttributeNode(attrClassBM)

	    Set attrStyleBM = document.createAttribute("style")
		attrStyleBM.value = "display:none;"
		banco_metalli.setAttributeNode(attrStyleBM)
		
	    Dim navbar_item_titoli: Set navbar_item_titoli = document.getElementById( "navbar_item_titoli" )
	    Dim titoli: Set titoli = document.getElementById( "titoli_section" )

	    Set attrClassTitoli = document.createAttribute("class")
		attrClassTitoli.value = "navbar_item navbar_item_selected"
		navbar_item_titoli.setAttributeNode(attrClassTitoli)

	    Set attrStyleTitoli = document.createAttribute("style")
		attrStyleTitoli.value = "display:block;"
		titoli.setAttributeNode(attrStyleTitoli)
	End Sub

	Sub select_tab(tab)
		Select case tab
			case "speds"
				select_tab_speds()
			Case "bm" 
				select_tab_bm()
				
			Case "titoli"
				select_tab_titoli()
			
		End Select
	End Sub
 