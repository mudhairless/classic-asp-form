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

            m_itemvalue = Request.Form(m_itemname)

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
                                                response.write "<!-- got: " & m_itemvalue & " value:" & fnmchild(1) &" -->"
                                            end if
                                        else
                                            fnResulti = fnmchild.validate()
                                        end if
                                    end if
                                    if(fnResulti) then exit for
                                next
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
                    else
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
        end sub

        private sub Class_Deinitialize
            for each mChild in m_formitems
                set mChild = nothing
            next
        end sub

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
            if(m_formitems.Exists(fnItemName)) then
                fnResult = Request.Form(fnItemName)
            end if
            Value = fnResult
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

    end class

%>
