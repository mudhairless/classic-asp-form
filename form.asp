<%

    'Classic ASP Form Engine

    const iTypeText = 0
    const iTypePassword = 1
    const iTypeCheckbox = 2
    const iTypeRadio = 3
    const iTypeFile = 4
    const iTypeHidden = 5
    const iTypeImage = 6
    const iTypeSearch = 5000
    const iTypeDateTime = 5001
    const iTypeDateTimeLocal = 5002
    const iTypeRange = 5003
    const iTypeEmail = 5004
    const iTypeDate = 5005
    const iTypeURL = 5006
    const iTypeMonth = 5007
    const iTypeColor = 5008
    const iTypeTel = 5009
    const iTypeWeek = 5010
    const iTypeNumber = 5011
    const iTypeTime = 5012
    const iTypeTextArea = 8000
    const iTypeSelect = 9000
    'const iTypeOption = 9001
    const iTypeOptGroup = 9002

	Private Function BStr2UStr(BStr)
		'Byte string to Unicode string conversion
		Dim lngLoop
		BStr2UStr = ""
		For lngLoop = 1 to LenB(BStr)
			BStr2UStr = BStr2UStr & Chr(AscB(MidB(BStr,lngLoop,1))) 
		Next
	End Function
	
	Private Function UStr2Bstr(UStr)
		'Unicode string to Byte string conversion
		Dim lngLoop
		Dim strChar
		UStr2Bstr = ""
		For lngLoop = 1 to Len(UStr)
			strChar = Mid(UStr, lngLoop, 1)
			UStr2Bstr = UStr2Bstr & ChrB(AscB(strChar))
		Next
	End Function 

	Private Function URLDecode(Expression)
		Dim strSource, strTemp, strResult
		Dim lngPos
		strSource = Replace(Expression, "+", " ")
		For lngPos = 1 To Len(strSource)
			strTemp = Mid(strSource, lngPos, 1)
			If strTemp = "%" Then
				If lngPos + 2 < Len(strSource) Then
					strResult = strResult & _
						Chr(CInt("&H" & Mid(strSource, lngPos + 1, 2)))
					lngPos = lngPos + 2
				End If
			Else
				strResult = strResult & strTemp
			End If
		Next
		URLDecode = strResult
	End Function


	Class FileItem
	
		Private m_strName
		Private m_strContentType
		Private m_strFileName
		Private m_Blob
		
		Public Property Get Name()
			Name = m_strName
		End Property
	  
		Public Property Let Name(vIn)
			m_strName = vIn
		End Property
	  
		Public Property Get ContentType()
			ContentType = m_strContentType
		End Property
	  
		Public Property Let ContentType(vIn)
			m_strContentType = vIn
		End Property

		public property Get Size()
			Size = lenb(m_blob)
		end property
	  
		Public Property Get FileName()
			FileName = m_strFileName
		End Property
	  
		Public Property Let FileName(vIn)
			m_strFileName = vIn
		End Property
	  
		Public Property Get Blob()
			Blob = m_Blob
		End Property
	  
		Public Property Let Blob(vIn)
			m_Blob = vIn
		End Property
	
		Public Sub Save(fnPath,fnUnique)
			Dim objFSO, objFSOFile
			Dim lngLoop, fnIsUnique, fnFinalPath, fnCnt
			Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
				fnFinalPath = objFSO.BuildPath(fnPath, m_strFileName)
				if(fnUnique) then
					fnIsUnique = objFSO.FileExists(server.mappath(fnFinalPath))
					fnCnt = 1
					while not fnIsUnique
						fnFinalPath = objFSO.BuildPath(fnPath, fnCnt & m_strFileName)
						fnIsUnique = objFSO.FileExists(server.mappath(fnFinalPath))
						fnCnt = fnCnt + 1
					wend
					m_strFileName = fnCnt & m_strFileName
				end if

				fnFinalPath = Server.MapPath(fnFinalPath)
				set objFSOFile = Server.CreateObject("ADODB.Stream")
					objFSOFile.mode = 3					
					objFSOFile.open
						for n = 1 to lenb(m_Blob)
							objFSOFile.writetext midb(m_Blob,n,1)
						next
						objFSOFile.Position = 2
						dim objBinfile
						set objBinfile = Server.CreateObject("ADODB.Stream")
							objBinFile.open
								objBinfile.type = 1
								objFSOFile.CopyTo objBinFile
								'objFSOFile.flush
								objBinFile.SaveToFile fnFinalpath,2
							objBinFile.close
						set objBinFile = nothing
					objFSOFile.close
				set objFSOFile = nothing
			set objFSO = nothing	
		End Sub
	End Class





'FormItem
    class FormItem
        private sub Class_Initialize
            m_children = 0
            set m_validation = Server.CreateObject("Scripting.Dictionary")
            set m_properties = Server.CreateObject("Scripting.Dictionary")
            m_validate = false
        end sub

        private sub Class_Deinitialize
            set m_validation = nothing
            set m_properties = nothing
        end sub

        public function validate()
			
            dim fnResulti
            fnResulti = true

            if(m_validate) then
                fnResulti = falses
                for each fnCheck in m_validation

                    select case fnCheck
                        case "required"
                            select case m_itemtype
	                            case iTypeRadio,iTypeSelect,iTypeOptGroup:
	                                for each fnmchild in m_formitems
	                                    if(not isEmpty(fnmchild)) then
	                                        if(isArray(fnmchild)) then
	                                            if(m_itemvalue = cstr(fnmchild(1))) then
	                                                fnResulti = true
	                                            else
	                                                fnResulti = false
	                                            end if
	                                        else
	                                            fnResulti = fnmchild.validate()
	                                        end if
	                                    end if
	                                    if(fnResulti) then exit for
	                                next
	                                
	                            case iTypeFile:
	                                fnResulti = isObject(m_itemvalue)
	                                
	                            case else:
	                                if(m_itemvalue <> "") then
	                                    fnResulti = true
	                                else
	                                    fnResulti = false
	                                end if
                            end select

                        case "regex"
                            if(m_itemvalue = "") then
                                fnResulti = false
                            else
                                set fntestRegex = new Regexp
                                    fntestRegex.Pattern = m_validation("regex")
                                    fnResulti = fntestRegex.Test(m_itemvalue)
                                set fntestRegex = nothing
                            end if
                            
                        case "content-type"
                        	if(m_itemtype = iTypeFile) then
                        		if(m_properties.Exists("accept")) then
                        			dim fnAccept,fnIpos
                        			fnAccept = m_properties.Item("accept")
                        			if(isObject(m_itemvalue)) then
                        				fnIpos = instr(fnAccept,m_itemvalue.ContentType)
                        				if(not isNull(fnIpos)) then
                        					if(fnIpos > 0) then
                        						fnResulti = true
                        					else
                        						fnResulti = false
                        					end if
                        				else
                        					fnResulti = false
                        				end if
                        			else
                        				fnResulti = false
                        			end if
                        		else
                        			fnResulti = false
                        		end if
                        	end if
                    end select

                next

                if(not fnResulti) then
                    if(m_validation.Exists("invalid-text")) then
                        m_valid = m_validation.Item("invalid-text")
                    else
                        m_valid = m_label & " is required."
                    end if
                end if

            end if
            validate = fnResulti

        end function

        public function add( fnArgumentArray )
            m_children = m_children + 1
            redim preserve m_formitems(m_children)
            select case m_itemtype
                case iTypeRadio:
                    m_formitems(m_children-1) = fnArgumentArray
                case iTypeSelect,iTypeOptGroup:
                    if(isArray(fnArgumentArray)) then
                        m_formItems(m_children-1) = fnArgumentArray
                    else
                        set m_formItems(m_children-1) = fnArgumentArray
                    end if
            end select
        end function

        public function display()

            dim fnResult
            fnResult = "<div class=""item " & m_class & """>"& vbCRLF & "<label for=""" & m_itemname & """>" & m_label & "</label>" & vbCRLF

            select case m_itemtype
                case iTypeTextArea:
                    fnResult = fnResult & "<textarea name=""" & m_itemname & """ "
                    if(m_itemdisabled) then
                        fnResult = fnResult & "disabled=""disabled"" "
                    end if

                    if(m_itemreadonly) then
                        fnResult = fnResult & "readonly=""readonly"" "
                    end if

                    fnResult = fnResult & "id=""" & m_itemname & """ "

                    fnResult = fnResult & ">"
                    if(m_itemvalue <> "") then
                        fnResult = fnResult & m_itemvalue
                    end if
                    fnResult = fnResult & "</textarea>" & vbCRLF

                case iTypeSelect:
                    fnResult = fnResult & "<select name=""" & m_itemname & """ "

                    fnResult = fnResult & "id=""" & m_itemname & """>" & vbCRLF

                    for each fnmchild in m_formitems
                        if(not isEmpty(fnmchild)) then
                            if(isArray(fnmchild)) then
                                if(m_itemvalue = cstr(fnmchild(1))) then
                                    selected = "selected=""selected"""
                                else
                                    selected = ""
                                end if
                                fnResult = fnResult & "<option value=""" & fnmchild(1) & """ " & selected & ">" & fnmchild(0) & "</option>" & vbCRLF
                            else
                                fnResult = fnResult & fnmchild.display()
                            end if
                        end if
                    next

                    fnResult = fnResult & "</select>" & vbCRLF

                case iTypeOptGroup:
                    fnResult = fnResult & "<optgroup label=""" & m_label & """>" & vbCRLF

                    for each fnmchild in m_formitems
                        if(not isEmpty(fnmchild)) then
                            if(m_itemvalue = cstr(fnmchild(1))) then
                                selected = "selected=""selected"""
                            else
                                selected = ""
                            end if
                            fnResult = fnResult & "<option value=""" & fnmchild(1) & """ "& selected &">" & fnmchild(0) & "</option>" & vbCRLF
                        end if
                    next

                    fnResult = fnResult & "</optgroup>" & vbCRLF

                case iTypeRadio:
                    for each fnmchild in m_formitems
                        if(not isEmpty(fnmchild)) then
                            if(m_itemvalue = cstr(fnmchild(1))) then
                                selected = "checked=""checked"""
                            else
                                selected = ""
                            end if
                            fnResult = fnResult & "<div class=""radioItem"">"&fnmchild(0) & vbCRLF
                            fnResult = fnResult & "<input type=""radio"" name="""& m_itemname & """ value=""" & fnmchild(1) & """ "& selected & "/>" & vbCRLF
                            fnResult = fnResult & "</div>" & vbCRLF
                        end if
                    next

                case else:
                    fnResult =  fnResult & "<input type=""" & typeToText() & """ " & _
                                "name=""" & m_itemname & """ " & _
                                "placeholder=""" & m_itemplaceholder & """ " & _
                                "tabindex=""" & m_itemtabindex & """ "
                    if(m_itemdisabled) then
                        fnResult = fnResult & "disabled=""disabled"" "
                    end if

                    if(m_itemreadonly) then
                        fnResult = fnResult & "readonly=""readonly"" "
                    end if

                    fnResult = fnResult & "id=""" & m_itemname & """ "

                    if(m_itemtype = iTypeCheckbox) then
                        if(m_itemvalue) then
                            fnResult = fnResult & "value=""True"" checked=""checked"" "
                        else
                            fnResult = fnResult & "value=""True"" "
                        end if
                    elseif(m_itemtype <> iTypeFile) then
                        fnResult = fnResult & "value=""" & m_itemvalue & """ "
                    end if

                    for each fnPropKey in m_properties
                        fnResult = fnResult & fnPropKey & "=""" & m_properties.Item(fnPropKey) & """ "
                    next

                    fnResult = fnResult & " />" & vbCRLF

            end select

            if(m_valid <> "") then
                fnResult = fnResult & "<div class=""validationMessage"">" & m_valid & "</div>" & vbCRLF
            end if
            fnResult = fnResult & "</div>" & vbCRLF

            display = fnResult

        end function

        private function typeToText()
            dim fnResult
            select case m_itemtype
                case iTypePassword:
                    fnResult = "password"
                case iTypeCheckbox:
                    fnResult = "checkbox"
                case iTypeRadio:
                    fnResult = "radio"
                case iTypeFile:
                    fnResult = "file"
                case iTypeHidden:
                    fnResult = "hidden"
                'case iTypeImage:
                '    fnResult = "image"
                case iTypeSearch:
                    fnResult = "search"
                case iTypeDateTime:
                    fnResult = "datetime"
                case iTypeDateTimeLocal:
                    fnResult = "datetimelocal"
                case iTypeRange:
                    fnResult = "range"
                case iTypeEmail:
                    fnResult = "email"
                case iTypeDate:
                    fnResult = "date"
                case iTypeURL:
                    fnResult = "url"
                case iTypeMonth:
                    fnResult = "month"
                case iTypeColor:
                    fnResult = "color"
                case iTypeTel:
                    fnResult = "tel"
                case iTypeWeek:
                    fnResult = "week"
                case iTypeNumber:
                    fnResult = "number"
                case iTypeTime:
                    fnResult = "time"
                case else:
                    fnResult = "text"
            end select
            typeToText = fnResult
        end function

        public property let Prop ( fnPname, fnVal )
            if(not m_properties.Exists(fnPname)) then
                m_properties.add fnPname, fnVal
            else
                m_properties(fnPname) = fnVal
            end if
        end property
        
        public property get Prop ( fnPname )
        		if(not m_properties.Exists(fnPname)) then
        			Prop = ""
        		else
        			Prop = m_properties.Item(fnPname)
        		end if
        	end property

        public property let Name ( fnIname )
            m_itemname = fnIname
        end property

        public property get Name
            Name = m_itemname
        end property

        public property let iType ( fnItype )
            m_itemtype = fnItype
        end property

        public property get iType
            iType = m_itemtype
        end property

        public property let Value( fnValue )
            m_itemvalue = fnValue
        end property

        public property get Value
            Value = m_itemvalue
        end property

        public property let Disabled( fnDisabled )
            m_itemdisabled = fnDisabled
        end property

        public property get Disabled
            Disabled = m_itemdisabled
        end property

        public property let ReadOnly( fnDisabled )
            m_itemreadonly = fnDisabled
        end property

        public property get ReadOnly
            ReadOnly = m_itemreadonly
        end property

        public property let TabIndex( fnTI )
            m_itemtabindex = fnTI
        end property

        public property get TabIndex
            TabIndex = m_itemtabindex
        end property

        public property let Label( fnLabel )
            m_label = fnLabel
        end property

        public property get Label
            Label = m_label
        end property

        public property let Placeholder( fnPH )
            m_itemplaceholder = fnPH
        end property

        public property get Placeholder
            Placeholder = m_itemplaceholder
        end property

        public property let iClass( fnC )
            m_class = fnC
        end property

        public property get iClass
            iClass = m_class
        end property

        public property let ValidateRequired( fnBool )
            if(not m_validation.Exists("required")) then
                m_validation.add "required", fnBool
            else
                m_validation.Item("required") = fnBool
            end if
            m_validate = true
        end property

        public property let ValidationFailedText( fnVFT )
            if(not m_validation.Exists("invalid-text")) then
                m_validation.add "invalid-text", fnVFT
            else
                m_validation.Item("invalid-text") = fnVFT
            end if
        end property

        public property let ValidateRegex( fnRegex )
            if(not m_validation.Exists("regex")) then
                m_validation.add "regex", fnRegex
            else
                m_validation.Item("regex") = fnRegex
            end if
            m_validate = true
        end property
        
        public property let ValidateContentType( fnBool )
        		if(not m_validation.Exists("content-type")) then
                m_validation.add "content-type", fnBool
            else
                m_validation.Item("content-type") = fnBool
            end if
            m_validate = true
        end property

        dim m_formitems()
        dim m_properties
        dim m_itemtype
        dim m_itemname
        dim m_itemvalue
        dim m_itemdisabled
        dim m_itemreadonly
        dim m_itemtabindex
        dim m_itemplaceholder
        dim m_children
        dim m_label
        dim m_class
        dim m_validation
        dim m_valid
        dim m_validate
    end class

    class Form

        private sub Class_Initialize
            m_children = 0
            set m_form_values = Server.CreateObject("Scripting.Dictionary")
            set m_form_files = Server.CreateObject("Scripting.Dictionary")
            ParseRequest
        end sub

        private sub Class_Deinitialize
            for each mChild in m_formitems
                set mChild = nothing
            next
            set m_form_values = nothing
            set m_form_files = nothing
        end sub

        	Private Sub ParseRequest()
			  Dim lngTotalBytes, lngPosBeg, lngPosEnd
			  Dim lngPosBoundary, lngPosTmp, lngPosFileName
			  Dim strBRequest, strBBoundary, strBContent
			  Dim strName, strFileName, strContentType
			  Dim strValue, strTemp
			  Dim objFile
							
			  'Grab the entire contents of the Request as a Byte string
			  lngTotalBytes = Request.TotalBytes
			  strBRequest = Request.BinaryRead(lngTotalBytes)
					
			  'Find the first Boundary
			  lngPosBeg = 1
			  lngPosEnd = _
			      InStrB(lngPosBeg, strBRequest, UStr2Bstr(Chr(13)))
			  If lngPosEnd > 0 Then
			    strBBoundary = _
			        MidB(strBRequest, lngPosBeg, lngPosEnd - lngPosBeg)
			    lngPosBoundary = InStrB(1, strBRequest, strBBoundary)
			  End If
			  If strBBoundary = "" Then
			  'The form must have been submitted *without* 
			  'ENCTYPE="multipart/form-data"
			  'But since we already called Request.BinaryRead,
			  'we can no longer access the Request.Form collection,
			  'so we need to parse the request and populate
			  'our own form collection.
			    lngPosBeg = 1
			    lngPosEnd = _
			        InStrB(lngPosBeg, strBRequest, UStr2BStr("&"))
			    Do While lngPosBeg < LenB(strBRequest)
			      'Parse the element and add it to the collection
			      strTemp = BStr2UStr(MidB(strBRequest, _
			          lngPosBeg, lngPosEnd - lngPosBeg))
			      lngPosTmp = InStr(1, strTemp, "=")
			      strName = URLDecode(Left(strTemp, lngPosTmp - 1))
			      strValue = URLDecode(Right(strTemp, _
			          Len(strTemp) - lngPosTmp))
			      m_form_values.Add strName, strValue
			      'Find the next element
			      lngPosBeg = lngPosEnd + 1
			      lngPosEnd = InStrB(lngPosBeg, _
			          strBRequest, UStr2BStr("&"))
			      If lngPosEnd = 0 Then
					lngPosEnd = LenB(strBRequest) + 1
				  End If
			    Loop
			  Else
			  'Form was submitted with ENCTYPE="multipart/form-data"
			  'Loop through all the boundaries, and parse them
			  'into either the Form or Files collections.
			  Do Until (lngPosBoundary = _
			      InStrB(strBRequest, strBBoundary & UStr2Bstr("--")))
			    'Get the element name
			    lngPosTmp = InStrB(lngPosBoundary, strBRequest, _
			        UStr2BStr("Content-Disposition"))
			    lngPosTmp = InStrB(lngPosTmp, _
			        strBRequest, UStr2BStr("name="))
			    lngPosBeg = lngPosTmp + 6
			    lngPosEnd = InStrB(lngPosBeg, _
			        strBRequest, UStr2BStr(Chr(34)))
			    strName = BStr2UStr(MidB(strBRequest, _
			        lngPosBeg, lngPosEnd - lngPosBeg))
			    'Look for an element named 'filename'
			    lngPosFileName = InStrB(lngPosBoundary, _
			        strBRequest, UStr2BStr("filename="))
			    'If found, we have a file, 
			    'otherwise it is a normal form element
			    If lngPosFileName <> 0 And lngPosFileName < _
			        InStrB(lngPosEnd, strBRequest, strBBoundary) Then
			      'It is a file. Get the FileName
			      lngPosBeg = lngPosFileName + 10
			      lngPosEnd = InStrB(lngPosBeg, _
			          strBRequest, UStr2BStr(chr(34)))
			      strFileName = BStr2UStr(MidB(strBRequest, _
			          lngPosBeg, lngPosEnd - lngPosBeg))
			      'Get the ContentType
			      lngPosTmp = InStrB(lngPosEnd, _
			          strBRequest, UStr2BStr("Content-Type:"))
			      lngPosBeg = lngPosTmp + 14
			      lngPosEnd = InstrB(lngPosBeg, _
			          strBRequest, UStr2BStr(chr(13)))
			      strContentType = BStr2UStr(MidB(strBRequest, _
			          lngPosBeg, lngPosEnd - lngPosBeg))
			      'Get the Content
			      lngPosBeg = lngPosEnd + 4
			      lngPosEnd = InStrB(lngPosBeg, _
			          strBRequest, strBBoundary) - 2
			      strBContent = MidB(strBRequest, _
			          lngPosBeg, lngPosEnd - lngPosBeg)
			      If strFileName <> "" And strBContent <> "" Then
			        'Create the File object, 
			        'and add it to the Files collection
			        Set objFile = New FileItem
			        objFile.Name = strName
			        objFile.FileName = Right(strFileName, _
			            Len(strFileName) - InStrRev(strFileName, "\"))
			        objFile.ContentType = strContentType
			        objFile.Blob = strBContent
			        m_form_files.Add strName, objFile
			      End If
			    Else 'It is a form element
			      'Get the value of the form element
			      lngPosTmp = InStrB(lngPosTmp, _
			          strBRequest, UStr2BStr(chr(13)))
			      lngPosBeg = lngPosTmp + 4
			      lngPosEnd = InStrB(lngPosBeg, _
			          strBRequest, strBBoundary) - 2
			      strValue =  _
			          BStr2UStr(MidB(strBRequest, _
			              lngPosBeg, lngPosEnd - lngPosBeg))
			      'Add the element to the collection
			      m_form_values.Add strName, strValue
			    End If
			    'Move to Next Element
			    lngPosBoundary = InStrB(lngPosBoundary + _
			        LenB(strBBoundary), strBRequest, strBBoundary)
			    Loop
			  End If
			End Sub

        public function display()

            dim fnResultf

            fnResultf = "<form name=""" & m_form_name & """ id=""" & m_form_id & """ enctype=""" & m_form_enctype & """ method=""" & m_form_method & """ accept-charset=""" & m_form_charsets & """ accept=""" & m_form_accept & """"

            if(m_form_action <> "") then
                fnResultf = fnResultf & " action=""" & m_form_action & """"
            else
                fnResultf = fnResultf & " action=""" & Request.ServerVariables("URL") & """"
            end if

            fnResultf = fnResultf & ">" & vbCRLF

            for each fnChild in m_formitems
                if(not isEmpty(fnChild)) then
                    fnResultf = fnResultf & fnChild.display()
                end if
            next

            if(m_form_reset <> "") then
                fnResultf = fnResultf & "<input type=""reset"" value=""" & m_form_reset & """/>" & vbCRLF
            end if
            fnResultf = fnResultf & "<input type=""submit"" value=""" & m_form_submit & """/>" & vbCRLF
            fnResultf = fnResultf & "</form>"

            display = fnResultf

        end function

        public sub add( fnFormItem )
            m_children = m_children + 1
            redim preserve m_formitems(m_children)

            
            if(fnFormItem.m_itemtype= iTypeFile) then
            	m_form_enctype = "multipart/form-data"
            	if(m_form_files.Exists(fnFormItem.m_itemname)) then
            		set fnFormItem.m_itemvalue = m_form_files.Item(fnFormItem.m_itemname)
            	end if
            else
            	if(m_form_values.Exists(fnFormItem.m_itemname)) then
            		fnFormItem.m_itemvalue = m_form_values.Item(fnFormItem.m_itemname)
            	end if
            end if
            
            set m_formitems(m_children-1) = fnFormItem
        end sub

        public function validate()
            fnResult = true

            for each mChild in m_formitems
                if(not isEmpty(mChild)) then
                    if(fnResult) then
                        fnResult = mChild.validate()
                    else
                        call mChild.validate()
                    end if
                end if
            next

            validate = fnResult
        end function

        public property get Value( fnItemName )
            dim fnResult
            fnResult = ""
            if(m_form_values.Exists(fnItemName)) then
                fnResult = m_form_values.Item(fnItemName)
            else
            	if(m_form_files.Exists(fnItemName)) then
            		set fnResult = m_form_files.Item(fnItemName)
            	end if
            end if
            if(isObject(fnResult)) then
            	set Value = fnResult
            else
            	Value = fnResult
            end if
        end property

        public property let Name( fnFormName )
            m_form_name = fnFormName
        end property

        public property let ID( fnFormID )
            m_form_id = fnFormID
        end property

        public property let Method( fnFormMethod )
            m_form_method = fnFormMethod
        end property

        public property let Encoding( fnFormEncType )
            m_form_enctype = fnformenctype
        end property

        public property let Action( fnFormAction )
            m_form_action = fnFormAction
        end property

        public property let AcceptCharset( fnFormAC )
            if(m_form_charsets = "") then
                m_form_charsets = fnFormAC
            else
                m_form_charsets = m_form_charsets & "," & fnFormAC
            end if
        end property

        public property let AcceptType( fnFormAT )
            if(m_form_accept = "") then
                m_form_accept = fnFormAT
            else
                m_form_accept = m_form_accept & "," & fnFormAT
            end if
        end property
        
        public property get Files ()
        		set Files = m_form_files
        end property

        public property let SubmitButtonText( fnST )
            m_form_submit = fnST
        end property

        public property let ResetButtonText( fnRT )
            m_form_reset = fnRT
        end property

        private m_formitems()
        private m_children
        private m_form_name
        private m_form_id
        private m_form_method
        private m_form_enctype
        private m_form_action
        private m_form_charsets
        private m_form_accept
        private m_form_submit
        private m_form_reset
        private m_form_values
        private m_form_files

    end class

%>