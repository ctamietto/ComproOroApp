	'rem versione 20230831_001
			
	' testata spedizione
	Class spedclass
    	Public kt,banco_metalli_id,banco_metalli_desc , data_ddt, numero_ddt
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
	
	Class CLparameter
		Public key,value
	End Class
	
	Class filterclass
		Public field,value
	End Class

	Public speddict: Set speddict = CreateObject("Scripting.Dictionary")
	
	Public speddetdict: Set speddetdict = CreateObject("Scripting.Dictionary")

	Public currspeddetdict: Set currspeddetdict = CreateObject("Scripting.Dictionary")

	Public bmdict: Set bmdict = CreateObject("Scripting.Dictionary")

	Public todict: Set todict = CreateObject("Scripting.Dictionary")
	
	Public spederrorsdict: Set spederrorsdict = CreateObject("Scripting.Dictionary")
	
	Public bmerrorsdict: Set bmerrorsdict = CreateObject("Scripting.Dictionary")
	
	Public toerrorsdict: Set toerrorsdict = CreateObject("Scripting.Dictionary")

	Public clparametersdict: Set clparametersdict = CreateObject("Scripting.Dictionary")

	Public filtersdict: Set filtersdict = CreateObject("Scripting.Dictionary")

	Function getCLParameter
		If ( InStr(objGestSped.commandLine,"mod=administrator") > 0) Then
			 Dim clpa : Set clpa = New CLparameter
			 clpa.key = "usermod"
			 clpa.value = "administrator"
			 clparametersdict.Add clpa.key,clpa
		Else 
			 Dim clpb : Set clpb = new CLparameter
			 clpb.key = "usermod"
			 clpb.value = "base"
			 clparametersdict.Add clpb.key,clpb
		End If 
	End Function

	Function isAdministrator
		Dim isa
		isa = False 
		If (clparametersdict.Exists("usermod")) Then 
			Set clp = clparametersdict.Item("usermod")
			If ( clp.value = "administrator" ) Then	
				isa = True 
			End If 
		End If 
		isAdministrator = isa
	End Function
	
	Function getLogicalDiskLetter
		getLogicalDiskLetter = Right(Left(objGestSped.commandLine,3),2)
	End Function 
	
	Function calcUniqueIdentifierForHash
		uifh = ""
		calcUniqueIdentifierForHash = uifh
	End Function
	
	Sub initialSizeAndPos
		window.moveTo 30, 30
	End Sub 

	Function BMErrorsStatus
		Dim ES
		ES = False  
		If Not IsNull(bmerrorsdict) And Not IsEmpty(bmerrorsdict) And bmerrorsdict.Count > 0 Then
			ES = True 
		End If 
		BMErrorsStatus = ES
	End Function

	Sub BMErrorsAdd(errcl)
		bmerrorsdict.Add errcl.field,errcl
	End Sub 

	Sub BMErrorsCleared
		bmerrorsdict.RemoveAll()	
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
		TitoliCalcTotaliSped(sped)
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
			sped.verga_fonderia  = CDbl(Replace(verga_fonderia.value,".",","))
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
			sped.titolo_fonderia  = CDbl(Replace(titolo_fonderia.value,".",","))
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
			sped.titolo_lab_controsaggio  = CDbl(Replace(titolo_lab_controsaggio.value,".",","))  
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
		
		TitoliSpedGet(sped)
		
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
	    Const ForReading = 1, ForWriting = 2, ForAppending = 8
    	Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0	
	
		filename = "bm.mydb"
		Set fso = CreateObject("Scripting.FileSystemObject")
		fullPathToFilename = fso.GetAbsolutePathName(filename)
		Rem alert(fullPathToFilename)
		Rem If (fso.FileExists(filename)) Then
		Rem	alert fullPathToFilename & " exists" 
		Rem End If 
		Rem restituisce un oggetto file stream 
		Set fbm = fso.OpenTextFile(fullPathToFilename, ForReading, True, TristateFalse)

		Dim bm: Set bm = New bmclass
		
		Do Until fbm.AtEndOfStream
      		bmrecord = fbm.ReadLine
      		Rem alert(bmrecord)      		
      		abm=Split(bmrecord,"!#!")
      		Rem alert(abm(0))
      		Rem alert(abm(1))

			Set bm = new bmclass
			With bm
				.kt = abm(0)
    			.desc = abm(1)
			End With
			bmdict.Add bm.kt, bm

	    Loop
	    fbm.Close
	
    End Sub

	Sub getStoreTO()
	    Const ForReading = 1, ForWriting = 2, ForAppending = 8
    	Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0	

		filename = "to.mydb"
		Set fso = CreateObject("Scripting.FileSystemObject")
		fullPathToFilename = fso.GetAbsolutePathName(filename)
		Set fto = fso.OpenTextFile(fullPathToFilename, ForReading, True, TristateFalse)

		Dim toi: Set toi = new toclass

		Do Until fto.AtEndOfStream
      		torecord = fto.ReadLine
      		ato=Split(torecord,"!#!")

			Set toi = new toclass
			With toi
				.kt = ato(0)
				tempdesc = Replace(ato(1),Chr(34),"")
    			.desc = tempdesc
    			.coefficiente =  CDbl(ato(2))
    			.coefficiente_titolo_stimato = ato(3)
			End With
			todict.Add toi.kt, toi

	    Loop
	    fto.Close

    End Sub

	Sub getStoredSpedDetts()	
		Const ForReading = 1, ForWriting = 2, ForAppending = 8
    	Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0	

		filename = "titoliddt.mydb"
		Set fso = CreateObject("Scripting.FileSystemObject")
		fullPathToFilename = fso.GetAbsolutePathName(filename)
		Set ftd = fso.OpenTextFile(fullPathToFilename, ForReading, True, TristateFalse)

		Dim speddet: Set speddet = new speddetclass

		Do Until ftd.AtEndOfStream
      		tdrecord = ftd.ReadLine
      		atd=Split(tdrecord,"!#!")

			Set speddet = new speddetclass
			With speddet
				.kt = atd(0)
				.fk = atd(1)
				.titolo_oro_id = atd(2)
	    		.grammi_lordi = CDbl(atd(3))
    			.grammi_puro_stimati = CDbl(atd(4)) 		
			End With
			speddetdict.Add speddet.kt, speddet

	    Loop
	    ftd.Close

    End Sub

	Rem rimuovi tutte le occorrenze del dettaglio di una spedizione ( legate ad una foreign key )
	Sub removeDettsOfSpedFK(fk)
		For Each i In speddetdict.Keys
			Set speddet = speddetdict.Item(i)
			If (speddet.fk = fk) Then
				speddetdict.Remove(i)
			End If
		Next 
    End Sub
    
    Sub storeCurDettsOfSped() 
		For Each i In currspeddetdict.Keys
			Set titolodett = currspeddetdict.Item(i)
			speddetdict.Add titolodett.kt, titolodett
		Next 
		currspeddetdict.RemoveAll()
    End Sub 

	Sub SpedDettTitoloCopy(ByRef DettTitoloDest,DettTitoloSrc)
		DettTitoloDest.kt = DettTitoloSrc.kt
		DettTitoloDest.fk = DettTitoloSrc.fk
		DettTitoloDest.titolo_oro_id = DettTitoloSrc.titolo_oro_id
		DettTitoloDest.titolo_oro_desc = DettTitoloSrc.titolo_oro_desc
		DettTitoloDest.grammi_lordi = DettTitoloSrc.grammi_lordi
		DettTitoloDest.grammi_puro_stimati = DettTitoloSrc.grammi_puro_stimati		
    End Sub

	Sub getDettsOfSped(fk)
		currspeddetdict.RemoveAll()
		If ( fk <> "" ) Then 
			For Each i In speddetdict.Keys
    			Set csddi = speddetdict.Item(i)
    			If ( csddi.fk = fk ) Then
    				Dim speddetcopy: Set speddetcopy = new speddetclass
					Call SpedDettTitoloCopy(speddetcopy,csddi)
    				currspeddetdict.Add speddetcopy.kt, speddetcopy
    			End If 
			Next
		End If      	
	End Sub 
	
	Sub getStoredSpeds()
		Const ForReading = 1, ForWriting = 2, ForAppending = 8
    	Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0	

		filename = "ddt.mydb"
		Set fso = CreateObject("Scripting.FileSystemObject")
		fullPathToFilename = fso.GetAbsolutePathName(filename)
		Set fsp = fso.OpenTextFile(fullPathToFilename, ForReading, True, TristateFalse)

		Dim sped: Set sped = new spedclass

		Do Until fsp.AtEndOfStream
      		sprecord = fsp.ReadLine
      		asp=Split(sprecord,"!#!")

			Set sped = new spedclass
			With sped
				.kt = asp(0)
    			.banco_metalli_id = asp(1)
    			.data_ddt = asp(2)
    			.numero_ddt = asp(3)
    			
    			.totale_grammi_rottami = CDbl(asp(4))
				.titolo_stimato_verga = CDbl(asp(5))
				.verga_stimata = CDbl(asp(6))		
				.totale_grammi_puro_stimato = CDbl(asp(7))
    			
    			.verga_fonderia = CDbl(asp(8))
    			.titolo_fonderia = CDbl(asp(9))
    			.totale_grammi_puro_fonderia = CDbl(asp(10))
    			
    			.titolo_lab_controsaggio = CDbl(asp(11))
				.puro_stimato_lab_controsaggio = CDbl(asp(12)) 
				
				.differenza_grammi = CDbl(asp(13))
				.differenza_percentuale = CDbl(asp(14))
				
				.differenza_grammi_con_saggio = CDbl(asp(15))
				.differenza_percentuale_con_saggio = CDbl(asp(16))
		
			End With
			speddict.Add sped.kt, sped

	    Loop
	    fsp.Close

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
    		    		    		
    		Set tdNodeTGR = document.createElement("td")
    		tdNodeTGR.innerHTML = "<p> " + CStr(sped.totale_grammi_rottami) + " </p> "
    		Set attrClassFieldTGR = document.createAttribute("class")
			attrClassFieldTGR.value = "spedizioni_field"
			tdNodeTGR.setAttributeNode(attrClassFieldTGR)
    		trNode.appendChild(tdNodeTGR)

    		Set tdNodeTSV = document.createElement("td")
    		tdNodeTSV.innerHTML = "<p> " + CStr(sped.titolo_stimato_verga) + " </p> "
    		Set attrClassFieldTSV = document.createAttribute("class")
			attrClassFieldTSV.value = "spedizioni_field"
			tdNodeTSV.setAttributeNode(attrClassFieldTSV)
    		trNode.appendChild(tdNodeTSV)

    		Set tdNodeVS = document.createElement("td")
    		tdNodeVS.innerHTML = "<p> " + CStr(sped.verga_stimata) + " </p> "
    		Set attrClassFieldVS = document.createAttribute("class")
			attrClassFieldVS.value = "spedizioni_field"
			tdNodeVS.setAttributeNode(attrClassFieldVS)
    		trNode.appendChild(tdNodeVS)
    		    		
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
    		tdNodeDiffGR.innerHTML = CStr(sped.differenza_grammi)
    		Set attrStyleGR = document.createAttribute("style")
			attrStyleGR.value = "background-color:#fcab69;color:white;font-weigth:bolder;text-align:center;"
			tdNodeDiffGR.setAttributeNode(attrStyleGR)
    		trNode.appendChild(tdNodeDiffGR)    		

    		Set tdNodeDiffPERC = document.createElement("td")
    		tdNodeDiffPERC.innerHTML = CStr(sped.differenza_percentuale) + "%"
    		Set attrStylePERC = document.createAttribute("style")
			attrStylePERC.value = "background-color:#fcab69;color:white;font-weigth:bolder;text-align:center;"
			tdNodeDiffPERC.setAttributeNode(attrStylePERC)
    		trNode.appendChild(tdNodeDiffPERC)    		

    		Set tdNodeDiffGRCS = document.createElement("td")
    		tdNodeDiffGRCS.innerHTML = CStr(sped.differenza_grammi_con_saggio)
    		Set attrStyleGRCS = document.createAttribute("style")
			attrStyleGRCS.value = "background-color:#fcb276;color:white;font-weigth:bolder;text-align:center;"
			tdNodeDiffGRCS.setAttributeNode(attrStyleGRCS)
    		trNode.appendChild(tdNodeDiffGRCS)    		

    		Set tdNodeDiffPERCCS = document.createElement("td")
    		tdNodeDiffPERCCS.innerHTML = CStr(sped.differenza_percentuale_con_saggio) + "%"
    		Set attrStylePERCCS = document.createAttribute("style")
			attrStylePERCCS.value = "background-color:#fcb276;color:white;font-weigth:bolder;text-align:center;"
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
        
    Sub TitoliCalcTotaliSped(sped)
    	sped.verga_stimata = 0
    	sped.titolo_stimato_verga = 0
    	sped.totale_grammi_rottami = 0
    	sped.totale_grammi_puro_stimato = 0
    	
    	For Each j In currspeddetdict.Keys
    		Set titolodett = currspeddetdict.Item(j)
    		Set titolo = todict.Item(titolodett.titolo_oro_id)
    		
    		sped.verga_stimata = sped.verga_stimata + ( titolodett.grammi_lordi * titolo.coefficiente_titolo_stimato )
    		sped.totale_grammi_rottami = sped.totale_grammi_rottami + titolodett.grammi_lordi
    		sped.totale_grammi_puro_stimato = sped.totale_grammi_puro_stimato + titolodett.grammi_puro_stimati
    	Next
    	If ( sped.totale_grammi_puro_stimato > 0 And sped.verga_stimata > 0 ) Then 
    		sped.titolo_stimato_verga = sped.totale_grammi_puro_stimato / sped.verga_stimata  * 1000 
    	End If 
    	
    	
    	sped.verga_stimata = Round(sped.verga_stimata,2)
    	sped.titolo_stimato_verga = Round(sped.titolo_stimato_verga,2)
    	sped.totale_grammi_rottami = Round(sped.totale_grammi_rottami,2)
    	sped.totale_grammi_puro_stimato = Round(sped.totale_grammi_puro_stimato,2)
    	
    End Sub
        
	Sub TitoliSpedGet(sped)
		For Each i In todict.Keys
    		Set titolo = todict.Item(i)
    		Set glel = document.getElementById( titolo.kt )
    		grammiLordo = 0
    		currentTitoloIndex = 0
    		foundCurrentTitoloIndex = false
    		
    		If (IsNull(glel.value) Or IsEmpty(glel.value) Or glel.value = "") Then
				grammiLordo = 0
			Else 
				grammiLordo  = CDbl(Replace(glel.value,".",","))  
			End If		

			For Each j In currspeddetdict.Keys
				Set titolodett = currspeddetdict.Item(j)
				If (titolodett.titolo_oro_id = titolo.kt) Then
					currentTitoloIndex = j
					foundCurrentTitoloIndex = True
				End If
			Next 

			If (foundCurrentTitoloIndex) Then 
				Set titolodettold = currspeddetdict.Item(currentTitoloIndex)
				If (grammiLordo > 0 ) Then 
					titolodettold.grammi_lordi = grammiLordo
					Rem todo calcola grammi_puro_stimati
					titolodettold.grammi_puro_stimati = Round(grammiLordo * titolo.coefficiente,2)
				Else 
					currspeddetdict.Remove(currentTitoloIndex)
				End If
			Else 
				If (grammiLordo > 0 ) Then 
					Dim titolodettnew: Set titolodettnew = new speddetclass	
					titolodettnew.kt = CreateGUID()
					titolodettnew.fk = sped.kt
					titolodettnew.titolo_oro_id = titolo.kt
					titolodettnew.titolo_oro_desc = titolo.desc
					titolodettnew.grammi_lordi = grammiLordo
					Rem todo calcola grammi_puro_stimati
					titolodettnew.grammi_puro_stimati = Round(grammiLordo * titolo.coefficiente,2)
					currspeddetdict.Add titolodettnew.kt, titolodettnew
				Else 
					Rem niente da fare 
				End If 
			End If     		
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
		inputGL.setAttribute "id",titolo.kt
		inputGL.setAttribute "type", "text"
		inputGL.setAttribute "onchange" , "SpedDetailValidate()"
		inputGL.setAttribute "onkeypress" , "return isDecimalKey(event)"
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

		Rem ricerca il titolo nell'elenco dei dettaglio corrente dei titolo
		For Each i In currspeddetdict.Keys
    		Set csd = currspeddetdict.Item(i)
			If (csd.titolo_oro_id = titolo.kt) Then
				inputGL.value = CStr(csd.grammi_lordi)
				inputGPS.value = CStr(csd.grammi_puro_stimati)
			End If 
		Next
	
    	tableNode.appendChild(trNode)
	End Sub

	Sub displayAllSpeds()
		For Each i In speddict.Keys
    		Set sped = speddict.Item(i)
    		getBMdesc(sped)
    		If (SpedRowFiltered(sped)) Then
    			'riga non visualizzata in quanto rispecchia i criteri del filtro
    		Else 
    			displaySped(sped)
    		End If 
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
		TitoliCalcTotaliSped(sped)
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
			
			Rem aggiungo i dettagli sui titoli della spedizione
			Rem prima li rimuovo tutti e poi li aggiungo
			If (isNewRecord) Then 
				Rem aggiungi la foreign key della spedizione ai dettagli   
				For Each i In currspeddetdict.Keys
					Set titolodett = currspeddetdict.Item(i)
					titolodett.fk = sped.kt
				Next 
			End If 
			
			' inserisci i dettagli titoli del ddt
			removeDettsOfSpedFK(sped.kt)
			storeCurDettsOfSped()
			
			' registra i dai nel file system
			SpedStoreInFile()
			SpedDetailStoreInFile()

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
		removeDettsOfSpedFK(kt)
		
		' registra i dai nel file system
		SpedStoreInFile()
		SpedDetailStoreInFile()
		
		Dim spedizioni_table: Set spedizioni_table = document.getElementById( "spedizioni_table" )
		clean_table(spedizioni_table)
		displayAllSpeds()
	End Sub

	Sub add_bm()
		Dim bm: Set bm = new bmclass
		BMErrorsCleared()
		bm.kt = ""
		bm.desc = ""
		BMDisplay(bm)
		show_mod_ins_bm()
	End Sub

	Sub modify_bm(kt)
		If bmdict.Exists(kt) Then
			BMErrorsCleared()
			Set bm = bmdict.Item(kt)
			BMDisplay(bm)
			show_mod_ins_bm()
		Else 
			MsgBox "Non esiste un banco metalli con chiave tecnica '" + kt + "'"
		End If

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
 
 	Sub displayBM(bm)
 	
 		Set tableNode = document.getElementById( "bm_table" )
    	Set trNode = document.createElement("tr")
    	Set attr = document.createAttribute("class")
		attr.value = "bmrow"
		trNode.setAttributeNode(attr)
    		
    	Set tdNodeDesc = document.createElement("td")
    	tdNodeDesc.innerHTML = "<p> " + CStr(bm.desc) + " </p> "    		
    	Set attrClassFieldDesc = document.createAttribute("class")
		attrClassFieldDesc.value = "bm_field"
		tdNodeDesc.setAttributeNode(attrClassFieldDesc)
    	trNode.appendChild(tdNodeDesc)
    	    	
    	Set iMNode = document.createElement("i")
    	Set attrClass = document.createAttribute("class")
		attrClass.value = "fa fa-sharp fa-solid fa-pencil icon_style"
		iMNode.setAttributeNode(attrClass)
    	Set attrOnClick = document.createAttribute("onClick")
		attrOnClick.value = "modify_bm('" + CStr(bm.kt) + "')"
		iMNode.setAttributeNode(attrOnClick)			
    	Set tdNodeMR = document.createElement("td")
    	tdNodeMR.appendChild(iMNode) 
    	trNode.appendChild(tdNodeMR) 

    	tableNode.appendChild(trNode)
 	End Sub
 
 
 	Sub displayAllBM()
		For Each i In bmdict.Keys
    		Set bm = bmdict.Item(i)
    		displayBM(bm)
		Next
    End Sub
    
     Sub displayAllTO()
		For Each i In todict.Keys
    		Set titolo = todict.Item(i)
    		displayTO(titolo)
		Next
    End Sub
    
    Sub displayTO(titolo)
 		Set tableNode = document.getElementById( "to_table" )
    	Set trNode = document.createElement("tr")
    	Set attr = document.createAttribute("class")
		attr.value = "torow"
		trNode.setAttributeNode(attr)
		
    	Set tdNodeDesc = document.createElement("td")
    	tdNodeDesc.innerHTML = "<p> " + CStr(titolo.desc) + " </p> "    		
    	Set attrClassFieldDesc = document.createAttribute("class")
		attrClassFieldDesc.value = "to_field"
		tdNodeDesc.setAttributeNode(attrClassFieldDesc)
    	trNode.appendChild(tdNodeDesc)
		
		Set tdNodeCoeff = document.createElement("td")
    	tdNodeCoeff.innerHTML = " <p> " + CStr(titolo.coefficiente) + " </p> "
    	Set attrClassFieldCoeff = document.createAttribute("class")
		attrClassFieldCoeff.value = "to_field_number"
		tdNodeCoeff.setAttributeNode(attrClassFieldCoeff)
    	trNode.appendChild(tdNodeCoeff)
		
		Set tdNodeCoeffTS = document.createElement("td")
    	tdNodeCoeffTS.innerHTML = " <p> " + CStr(titolo.coefficiente_titolo_stimato) + " </p> "
    	Set attrClassFieldCoeffTS = document.createAttribute("class")
		attrClassFieldCoeffTS.value = "to_field_number"
		tdNodeCoeffTS.setAttributeNode(attrClassFieldCoeffTS)
    	trNode.appendChild(tdNodeCoeffTS)

    	Set iMNode = document.createElement("i")
    	Set attrClass = document.createAttribute("class")
		attrClass.value = "fa fa-sharp fa-solid fa-pencil icon_style"
		iMNode.setAttributeNode(attrClass)
    	Set attrOnClick = document.createAttribute("onClick")
		attrOnClick.value = "modify_titolo('" + CStr(titolo.kt) + "')"
		iMNode.setAttributeNode(attrOnClick)			
    	Set tdNodeMR = document.createElement("td")
    	tdNodeMR.appendChild(iMNode) 
    	trNode.appendChild(tdNodeMR) 
		
    	tableNode.appendChild(trNode)
		
    End Sub

	Sub show_bm_list() 
    	document.getElementById("list_banco_metalli_container").style.display="block"
    	document.getElementById("mod_add_banco_metalli_container").style.display="none"
    End Sub
	
	Sub show_mod_ins_bm() 
    	document.getElementById("list_banco_metalli_container").style.display="none"
    	document.getElementById("mod_add_banco_metalli_container").style.display="block"
    End Sub

	Sub BMDetailCopy(ByRef bmDest,bmSrc)
		bmDest.kt = bmSrc.kt		
		bmDest.desc = bmSrc.desc
    End Sub



	Sub BMDetailValidate()
  		Dim existErrors
		existErrors = False 
		BMErrorsCleared()
			
		Dim bm: Set bm = new bmclass
		Dim isNewRecord
		isNewRecord = False
				
		Call BMDetailGet(bm,isNewRecord)
		BMDisplay(bm)
    End Sub

	Sub submit_bm
  		Dim existErrors
		existErrors = False 
		BMErrorsCleared()

		Dim bm: Set bm = new bmclass
		Dim isNewRecord
		isNewRecord = False
		
		Call BMDetailGet(bm,isNewRecord)
		existErrors = BMErrorsStatus()
		
		If (Not existErrors) Then
			If (Not isNewRecord) Then 
				If bmdict.Exists(bm.kt) Then
					Set bmDest = bmdict.Item(bm.kt)
					Call BMDetailCopy(bmDest,bm)
				Else 
					MsgBox "Non esiste una banco metalli con chiave tecnica '" + kt + "'"
				End If
			Else
				bm.kt = CreateGUID()
 				bmdict.Add bm.kt, bm
			End If 
		
			BMStoreInFile()
		
			Dim bm_table: Set bm_table = document.getElementById( "bm_table" )
			clean_table(bm_table)
			displayAllBM()
			show_bm_list()

			Rem aggiorna la ragione sociale di tutti i ddt 
			Dim spedizioni_table: Set spedizioni_table = document.getElementById( "spedizioni_table" )
			clean_table(spedizioni_table)
			displayAllSpeds()
			
			'ricostruisci la select multi del filtro banco metalli
			BuildMultiSelectBM()
			RebuildFilterMultiSelectBM()
		Else 
			BMDisplay(bm)
			MsgBox "ci sono errori"
		End If 

	End Sub
	
	
	Sub BMDisplay(bm)
		Set ktbm = document.getElementById( "ktbm" )
		ktbm.value = bm.kt

		Set bm_ragione_sociale = document.getElementById( "bm_ragione_sociale" )
		bm_ragione_sociale.value = bm.desc
		
		Set bm_ragione_sociale_error_list = document.getElementsByName("bm_ragione_sociale_error") 
 		For Each Elem In bm_ragione_sociale_error_list
 			Dim bm_ragione_sociale_error_object: Set bm_ragione_sociale_error_object = New errorclass			
 			If bmerrorsdict.Exists("desc") Then
				Set bm_ragione_sociale_error_object = bmerrorsdict.Item("desc")
		  		Elem.innerHTML = bm_ragione_sociale_error_object.desc
		  	Else 
		  		Elem.innerHTML = ""
			End if 
 		Next

	End Sub 
	
	
	Sub BMDetailGet(ByRef bm,ByRef isNewRecord)
		BMErrorsCleared()
		isNewRecord = False 
		
		Set kt = document.getElementById( "ktbm" )
		If (IsNull(kt.value) Or IsEmpty(kt.value)) Then
			bm.kt = ""
		Else 
			bm.kt = kt.value
		End If		
		
		Set bm_ragione_sociale = document.getElementById( "bm_ragione_sociale" )

		If (IsNull(bm_ragione_sociale.value) Or IsEmpty(bm_ragione_sociale.value)) Then
			bm.desc = ""
		Else 
			bm.desc  = bm_ragione_sociale.value
		End If		
		
		If (bm.desc = "") Then
			Dim errclDesc: Set errclDesc = New errorclass
			errclDesc.cod   = "000001"
			errclDesc.tipo  = "REQUIRED"
			errclDesc.field = "desc"
			errclDesc.desc  = "IMPOSTARE RAGIONE SOCIALE"
			BMErrorsAdd(errclDesc)
		End If
		
		If (bm.kt = "") Then
			isNewRecord = True
		End If
					
    End Sub


	Sub BMStoreInFile()
		filename = "bm.mydb"
	    Const ForReading = 1, ForWriting = 2, ForAppending = 8
    	Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
    
    	Dim fs, f
    	Set fs = CreateObject("Scripting.FileSystemObject")
    	Set f = fs.OpenTextFile(filename, ForWriting, True, TristateFalse)
    	For Each i In bmdict.Keys
    		Set bm = bmdict.Item(i)
    		Rem componi la stringa che corrisponde al record  
    		bmrecord = bm.kt & "!#!" & bm.desc
	    	f.WriteLine bmrecord
		Next 
    	f.Close
	End Sub
	
	Sub show_to_list() 
    	document.getElementById("list_titoli_container").style.display="block"
    	document.getElementById("mod_add_titoli_container").style.display="none"
    End Sub

	Sub show_mod_ins_to() 
    	document.getElementById("list_titoli_container").style.display="none"
    	document.getElementById("mod_add_titoli_container").style.display="block"
    End Sub
    
    Sub TODetailGet(toi,isNewRecord)
		TOErrorsCleared()
		isNewRecord = False 
		
		Set kt = document.getElementById( "ktto" )
		If (IsNull(kt.value) Or IsEmpty(kt.value)) Then
			toi.kt = ""
		Else 
			toi.kt = kt.value
		End If		

		Set to_desc = document.getElementById( "to_desc" )
		

		If (IsNull(to_desc.value) Or IsEmpty(to_desc.value)) Then
			toi.desc = ""
		Else 
			toi.desc  = Replace(to_desc.value,".",",")
		End If		
		
		If (toi.desc = "") Then
			Dim errclDesc: Set errclDesc = New errorclass
			errclDesc.cod   = "000001"
			errclDesc.tipo  = "REQUIRED"
			errclDesc.field = "desc"
			errclDesc.desc  = "IMPOSTARE DESCRIZIONE"
			TOErrorsAdd(errclDesc)
		End If


		Set coefficiente = document.getElementById( "coefficiente" )
		
		If (IsNull(coefficiente.value) Or IsEmpty(coefficiente.value) Or (Len(coefficiente.value) = 0 )) Then
			toi.coefficiente = 0
		Else 
			toi.coefficiente  = CDbl(Replace(coefficiente.value,".",","))
		End If		

		If (toi.coefficiente <= 0) Then
			Dim errclC: Set errclC = New errorclass
			errclC.cod   = "000002"
			errclC.tipo  = "GREATER_THEN"
			errclC.field = "coefficiente"
			errclC.desc  = "IMPOSTARE COEFFICIENTE"
			TOErrorsAdd(errclC)
			anyErrors = True  
		End If

		Set coefficiente_titolo_stimato = document.getElementById( "coefficiente_titolo_stimato" )
		
		If (IsNull(coefficiente_titolo_stimato.value) Or IsEmpty(coefficiente_titolo_stimato.value) Or (Len(coefficiente_titolo_stimato.value) = 0 )) Then
			toi.coefficiente_titolo_stimato = 0
		Else 
			toi.coefficiente_titolo_stimato  = CDbl(Replace(coefficiente_titolo_stimato.value,".",","))
		End If		

		If (toi.coefficiente_titolo_stimato <= 0) Then
			Dim errclCTS: Set errclCTS = New errorclass
			errclCTS.cod   = "000003"
			errclCTS.tipo  = "GREATER_THEN"
			errclCTS.field = "coefficiente_titolo_stimato"
			errclCTS.desc  = "IMPOSTARE COEFFICIENTE TITOLO STIMATO"
			TOErrorsAdd(errclCTS)
			anyErrors = True  
		End If
		
		If (toi.kt = "") Then
			isNewRecord = True
		End If

    End Sub 
    
    Sub TODetailCopy(toDest,toSrc)
		toDest.kt = toSrc.kt		
		toDest.desc = toSrc.desc
		toDest.coefficiente = toSrc.coefficiente
		toDest.coefficiente_titolo_stimato = toSrc.coefficiente_titolo_stimato
    End Sub
    
    Sub TOStoreInFile
		filename = "to.mydb"
	    Const ForReading = 1, ForWriting = 2, ForAppending = 8
    	Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
    
    	Dim fs, f
    	Set fs = CreateObject("Scripting.FileSystemObject")
    	Set f = fs.OpenTextFile(filename, ForWriting, True, TristateFalse)
    	For Each i In todict.Keys
    		Set toi = todict.Item(i)
    		Rem componi la stringa che corrisponde al record  
    		torecord = toi.kt & "!#!" & Chr(34) & toi.desc & Chr(34) & "!#!" & CStr(toi.coefficiente) & "!#!" & CStr(toi.coefficiente_titolo_stimato)
	    	f.WriteLine torecord
		Next 
    	f.Close
    End Sub
     
    Sub submit_to
  		Dim existErrors
		existErrors = False 
		TOErrorsCleared()

		Dim toi: Set toi = new toclass
		Dim isNewRecord
		isNewRecord = False
		
		Call TODetailGet(toi,isNewRecord)
		existErrors = TOErrorsStatus()
				
		If (Not existErrors) Then
			If (Not isNewRecord) Then 
				If todict.Exists(toi.kt) Then
					Set toDest = todict.Item(toi.kt)
					Call TODetailCopy(toDest,toi)
				Else 
					MsgBox "Non esiste un titolo con chiave tecnica '" + kt + "'"
				End If
			Else
				toi.kt = CreateGUID()
 				todict.Add toi.kt, toi
			End If 
		
			TOStoreInFile()
		
			Dim to_table: Set to_table = document.getElementById( "to_table" )
			clean_table(to_table)
			displayAllTO()
			show_to_list()

			Rem aggiorna il titolo di tutti i ddt 
			Dim spedizioni_table: Set spedizioni_table = document.getElementById( "spedizioni_table" )
			clean_table(spedizioni_table)
			displayAllSpeds()
						
		Else 
			TODisplay(toi)
			MsgBox "ci sono errori"
		End If 

    End Sub 
    
    Sub TODetailValidate(element,isDecimal)
  		Dim existErrors
		existErrors = False 
		TOErrorsCleared()
			
		Dim toi: Set toi = new toclass
		Dim isNewRecord
		isNewRecord = False
				
		Call TODetailGet(toi,isNewRecord)
		TODisplay(toi)
    End Sub 
    
    Sub TOErrorsCleared
		toerrorsdict.RemoveAll()	
	End Sub

	Sub TODisplay(toi)
		Set ktto = document.getElementById( "ktto" )
		ktto.value = toi.kt

		Set to_desc = document.getElementById( "to_desc" )
		to_desc.value = toi.desc
		
		
		Set to_desc_error_list = document.getElementsByName("to_desc_error") 
 		For Each Elem In to_desc_error_list
 			Dim to_desc_error_list_object: Set to_desc_error_list_object = New errorclass			
 			If toerrorsdict.Exists("desc") Then
				Set to_desc_error_list_object = toerrorsdict.Item("desc")
		  		Elem.innerHTML = to_desc_error_list_object.desc
		  	Else 
		  		Elem.innerHTML = ""
			End if 
 		Next

		Set coefficiente = document.getElementById( "coefficiente" )
		coefficiente.value = CStr(toi.coefficiente)
		
		Set coefficiente_error_list = document.getElementsByName("coefficiente_error") 
 		For Each Elem In coefficiente_error_list
 			Dim coefficiente_error_list_object: Set coefficiente_error_list_object = New errorclass			
 			If toerrorsdict.Exists("coefficiente") Then
				Set coefficiente_error_list_object = toerrorsdict.Item("coefficiente")
		  		Elem.innerHTML = coefficiente_error_list_object.desc
		  	Else 
		  		Elem.innerHTML = ""
			End if 
 		Next

		Set coefficiente_titolo_stimato = document.getElementById( "coefficiente_titolo_stimato" )
		coefficiente_titolo_stimato.value = CStr(toi.coefficiente_titolo_stimato)

		Set coefficiente_titolo_stimato_error_list = document.getElementsByName("coefficiente_titolo_stimato_error") 
 		For Each Elem In coefficiente_titolo_stimato_error_list
 			Dim coefficiente_titolo_stimato_error_list_object: Set coefficiente_titolo_stimato_error_list_object = New errorclass			
 			If toerrorsdict.Exists("coefficiente_titolo_stimato") Then
				Set coefficiente_titolo_stimato_error_list_object = toerrorsdict.Item("coefficiente_titolo_stimato")
		  		Elem.innerHTML = coefficiente_titolo_stimato_error_list_object.desc
		  	Else 
		  		Elem.innerHTML = ""
			End if 
 		Next

	End Sub
	
	Sub add_to()
		Dim toi: Set toi = new toclass
		TOErrorsCleared()
		toi.kt = ""
		toi.desc = ""
		toi.coefficiente = 0
		toi.coefficiente_titolo_stimato = 0
		TODisplay(toi)
		show_mod_ins_to()
	End Sub

	Sub modify_titolo(kt)
		If todict.Exists(kt) Then
			TOErrorsCleared()
			Set toi = todict.Item(kt)
			TODisplay(toi)
			show_mod_ins_to()
		Else 
			MsgBox "Non esiste un titolo con chiave tecnica '" + kt + "'"
		End If
	End Sub

	Function TOErrorsStatus
		Dim ES
		ES = False  
		If Not IsNull(toerrorsdict) And Not IsEmpty(toerrorsdict) And toerrorsdict.Count > 0 Then
			ES = True 
		End If 
		TOErrorsStatus = ES
	End Function

	Sub TOErrorsAdd(errcl)
		toerrorsdict.Add errcl.field,errcl
	End Sub 

	Sub getBMdesc(sped)
		If bmdict.Exists(sped.banco_metalli_id) Then
			Set bm = bmdict.Item(sped.banco_metalli_id)
			sped.banco_metalli_desc = bm.desc
		Else 
			sped.banco_metalli_desc = ""
		End If
	End Sub
	
	
	Sub SpedDetailStoreInFile
		filename = "titoliddt.mydb"
	    Const ForReading = 1, ForWriting = 2, ForAppending = 8
    	Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
    
    	Dim fs, f
    	Set fs = CreateObject("Scripting.FileSystemObject")
    	Set f = fs.OpenTextFile(filename, ForWriting, True, TristateFalse)
    	For Each i In speddetdict.Keys
    		Set speddet = speddetdict.Item(i)
    		Rem componi la stringa che corrisponde al record  
    		speddetrecord = speddet.kt & "!#!" & speddet.fk & "!#!" & speddet.titolo_oro_id & "!#!" & CStr(speddet.grammi_lordi) & "!#!" & CStr(speddet.grammi_puro_stimati)
	    	f.WriteLine speddetrecord
		Next 
    	f.Close
    End Sub
    

	Sub SpedStoreInFile
		filename = "ddt.mydb"
	    Const ForReading = 1, ForWriting = 2, ForAppending = 8
    	Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
    
    	Dim fs, f
    	Set fs = CreateObject("Scripting.FileSystemObject")
    	Set f = fs.OpenTextFile(filename, ForWriting, True, TristateFalse)
    	For Each i In speddict.Keys
    		Set sped = speddict.Item(i)
    		Rem componi la stringa che corrisponde al record  
    		spedrecord = sped.kt & "!#!" & sped.banco_metalli_id & "!#!" & sped.data_ddt & "!#!" & sped.numero_ddt & "!#!" & CStr(sped.totale_grammi_rottami) & "!#!" & CStr(sped.titolo_stimato_verga) & "!#!" & CStr(sped.verga_stimata) & "!#!" & CStr(sped.totale_grammi_puro_stimato) & "!#!" & CStr(sped.verga_fonderia) & "!#!" & CStr(sped.titolo_fonderia) & "!#!" & CStr(sped.totale_grammi_puro_fonderia) & "!#!" & CStr(sped.titolo_lab_controsaggio) & "!#!" & CStr(sped.puro_stimato_lab_controsaggio) & "!#!" & CStr(sped.differenza_grammi) & "!#!" & CStr(sped.differenza_percentuale) & "!#!" & CStr(sped.differenza_grammi_con_saggio) & "!#!" & CStr(sped.differenza_percentuale_con_saggio)
	    	f.WriteLine spedrecord
		Next 
    	f.Close
    End Sub
    
 	Sub submit_search
 		Dim filteritem : Set filteritem = New filterclass
		filteritem.field = "search_bm"
		filteritem.value = "" 
 		filterfound = False 
 	
 	    Dim search_bm: Set search_bm = document.getElementById("search_bm")
 	    For Each optioni In search_bm.options.all
 	    	If Not IsNull(optioni.value) And Not IsEmpty(optioni.value) And Len(optioni.value) > 0 Then
 	    		If Not IsNull(optioni.selected) And Not IsEmpty(optioni.selected) And optioni.selected Then
					filteritem.value = filteritem.value + optioni.value
					filterfound = True
 	    		End If 
 	    	End If 
 	    Next
 	    If filterfound Then
 	    	If filtersdict.Exists(filteritem.field) Then 
				filtersdict.Remove(filteritem.field) 
 	    	End If 
 	    	filtersdict.Add filteritem.field,filteritem
 	    Else 
 	    	'ripulisci eventuali filtri di search_bm derivanti da precedenti ricerche
 	    	filtersdict.Remove("search_bm")
 	    End If
		
		Dim spedizioni_table: Set spedizioni_table = document.getElementById( "spedizioni_table" )
		clean_table(spedizioni_table)
		displayAllSpeds()
 	End Sub
 	
 	Sub BuildMultiSelectBM
 	    	Dim search_bm: Set search_bm = document.getElementById("search_bm")
    		BuildSelectBM(search_bm)
    End Sub
        
	Function FiltersStatus
		Dim ES
		ES = False  
		If Not IsNull(filtersdict) And Not IsEmpty(filtersdict) And filtersdict.Count > 0 Then
			ES = True 
		End If 
		FiltersStatus = ES
	End Function
	
	Sub FilterAdd(filteri)
		filtersdict.Add filteri.field,filteri
	End Sub 

    Sub FiltersCleared
		filtersdict.RemoveAll()	
	End Sub
		
	Sub FiltersDisplayCleared
		If (FiltersStatus) Then
			For Each filteri In filtersdict
				If (filteri = "search_bm") Then
 		    		Dim search_bm: Set search_bm = document.getElementById("search_bm")
	    			BuildSelectBM(search_bm)		 
				End If 
			Next
			filtersdict.RemoveAll()
		Else 	
 		    Dim search_bm_2: Set search_bm_2 = document.getElementById("search_bm")
	    	BuildSelectBM(search_bm_2)		 
		End If
		
		Dim spedizioni_table: Set spedizioni_table = document.getElementById( "spedizioni_table" )
		clean_table(spedizioni_table)
		displayAllSpeds()
		
	End Sub
		
	Function SpedRowFiltered(sped)
		Dim srf
		srf = False 
		If FiltersStatus() Then
			For Each filteri In filtersdict
				If (filteri = "search_bm") Then
					Set filtero = filtersdict.Item(filteri)
					If InStr(filtero.value,sped.banco_metalli_id) = 0 Then
						srf = True  
					End If 
				End If 
			Next
		End If 
		SpedRowFiltered = srf
	End Function

	Sub RebuildFilterMultiSelectBM
		Dim search_bm: Set search_bm = document.getElementById("search_bm")
		If FiltersStatus() Then
			For Each filteri In filtersdict
				If (filteri = "search_bm") Then
					Set filtero = filtersdict.Item(filteri)
					For Each optioni In search_bm.options.all
 	    				If Not IsNull(optioni.value) And Not IsEmpty(optioni.value) And Len(optioni.value) > 0 Then
 	    					If (filtero.value = optioni.value) Then
 	    						 optioni.selected = True 
 	    					End If 
 	    				End If 
					Next
				End If 
			Next 		
		End If 
	End Sub 