<!-- #include file="form.asp" -->
<%
    isValid = false
    set myForm = new Form

        set nameItem = new FormItem
        with nameItem
            .iType = iTypeText
            .Name = "userName"
            .Placeholder = "User name"
            .Label = "User name"
            .ValidateRegex = "^[a-z0-9_-]{3,15}$"
            .ValidationFailedText = "Username must be between 3 and 15 characters, letters, numbers, underscore and dash allowed."
        end with

        set passItem = new FormItem
        with passItem
            .iType = iTypePassword
            .Name = "userPassword"
            .Placeholder = "Password"
            .Label = "Password"
            .ValidateRequired = true
        end with

        set checkItem = new FormItem
        with checkItem
            .itype = iTypeCheckbox
            .Name = "login"
            .Label = "Stay logged in"
        end with

        set textItem = new FormItem
        with textItem
            .iType = iTypeTextArea
            .Name = "comments"
            .Label = "Your Comments"
            .ValidateRequired = true
        end with

        set listItem = new FormItem
        with listItem
            .iType = iTypeSelect
            .Name = "favCheese"
            .Label = "Favorite Cheese"
            .add array("Cheddar",1)
            .add array("American",2)
            .add array("Swiss",3)
            .add array("Muenster",4)
            .add array("Brie",5)
        end with

        set radioItem = new FormItem
        with radioItem
            .iType = iTypeRadio
            .Name = "likeWaffles"
            .Label = "Do you like waffles?"
            .add array("Yes",1)
            .add array("No",0)
        end with

        set rangeItem = new Formitem
        with rangeItem
            .iType = iTypeRange
            .Name = "numOfWaffles"
            .Label = "How many waffles do you want?"
            .Prop("min") = 0
            .Prop("max") = 20
            .Prop("step") = 1
        end with

        set colorItem = new FormItem
        with colorItem
            .iType = iTypeColor
            .Name = "favcolor"
            .Label = "Favorite Color"
        end with

        set avatarItem = new FormItem
        with avatarItem
        	.itype = iTypeFile
        	.Name = "avatar"
        	.Label = "New Avatar"
        	.Prop("accept") = "image/png,image/jpeg"
        	.ValidateContentType = true
        	.ValidationFailedText = "Invalid file type, must be png or jpg."
        end with

        with myForm
            .Action = "/formtest"
            .Method = "post"
            .SubmitButtonText = "Login"
            .ResetButtonText = "Cancel"

            .add nameItem
            .add passItem
            .add checkItem
            .add textItem
            .add listItem
            .add radioItem
            .add rangeItem
            .add colorItem
            .add avatarItem
        end with

        if(Request.TotalBytes > 0) then
            isValid = myForm.validate()
            if(isValid) then
            	'save the file to disk
            	wassaved = myForm.Value("avatar").Save("/uploads", iUploadUnique)
            	'uncomment to output the uploaded image
            	'response.clear
            	'response.AddHeader "Content-Type",myForm.Value("avatar").ContentType
            	'response.binarywrite myForm.Value("avatar").Blob
            	'response.end
            end if
        end if

        %>
<!DOCTYPE html>
<html>
<head>
    <title>Form Test</title>
    <style type="text/css">
        .validationMessage {
            display: inline;
        }
    </style>
</head>
<body>
    <%=myForm.display()%>
    <p>
    	<%if(isObject(myForm.Value("avatar"))) then%>
    	You uploaded: <%=myForm.Value("avatar").Filename%>
    	<%end if%>
    </p>
</body>
</html>
<%
    set myForm = nothing
%>