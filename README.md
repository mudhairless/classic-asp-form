Simple Forms for Classic ASP
===========================

Simple Forms makes building and verification of forms easy in Classic ASP.
It is important that you do not use the *Request.Form* object at the same
time as Simple Forms as using one will break the other (ASP limitation).
Using Simple Forms to read forms is just as easy as using *Request.Form*, although
the api is different it is very similar to *Request.Form*.

Example
-------

For an example of usage see the included file: example-form.asp

Downloads
---------

The most current version can always be downloaded from:
[Github](https://github.com/mudhairless/classic-asp-form/archive/master.zip)

Class Reference
---------------

Form
----

  * add ( *FormItem* )

    > Adds the requested *FormItem* to the form. This is not necessary for basic usage, only needed if using *display* or *validate*.

  * display ()

    > Returns a string containing the HTML necessary to render the specified form items. Each item is wrapped in an enclosing div and if label text is specified has a label element as well.

  * validate ()

    > Tests elements marked for validation returning true if all tests passed or false otherwise. If validation has been performed before displaying the form then invalid items will have a message appended to them.

  * Value( formItemName ) *GET*

    > Returns the value of the form element that was submitted. This value will be a string for all element types except for the File type which will return a *FileItem* object.

  * Files () *GET*

    > Returns a Dictionary containing all of the form's submitted files only. The key in the Dictionary is the form element's name and the value is a *FileItem* object.

  * Name ( string ) *SET*

    > Sets the name property of the form element.

  * ID ( string ) *SET*

    > Sets the ID property of the form element.

  * Method ( string ) *SET*

    > Sets the method the browser will use to submit the form, defaults to "GET", "POST" required for file upload.

  * Encoding ( string ) *SET*

    > Sets the encoding the browser will use to encode text before submitting it.

  * Action ( url_string ) *SET*

    > Sets the url the browser will submit the form to, defaults to the current script.

  * AcceptCharset ( csv_string ) *SET*

    > Tell the browser what character sets you are capable of handling.

  * AcceptType ( csv_string ) *SET*

    > Tell the browser what types you will accept.

  * SubmitButtonText ( string ) *SET*

    > What the submit button will say, defaults to "Submit"

  * ResetButtonText ( string ) *SET*

    > If provided will include a form reset button with the requested label.


FormItem
--------

  * add ( *array()* )

    > This function is used for items that have children or are grouped like radio buttons, selects and optgroups. There are two ways to call this function depending on the item type:

      * Radio, Select, OptGroup

        > add( array( "item label", value ) )

          Will add a new item in the parent item.

      * Select

        > add( array( optgroup_FormItem ) )

  * Name ( string ) *GET/SET*

    > Sets the name property of the element, will be used as label if none provided. ID of element will match this value.

  * iType ( integer_constant ) *GET/SET*

    > Sets the type of this element.

  * Value ( mixed ) *GET/SET*

    > Sets a default value for the item. Value is filled in when the item is added to the form object so any submitted values will overwrite the default value.

  * Prop ( propname, value ) *GET/SET*

    > Allows you to set any other key/value pair on the item. See example for recommended usage.

  * Disabled ( bool ) *GET/SET*

    > Do not display the form element or submit it's value.

  * Readonly ( bool ) *GET/SET*

    > Display the form element but do not allow it's value to change.

  * TabIndex ( integer ) *GET/SET*

    > Manually set the order in which the TAB key will move through the form.

  * Label ( string ) *GET/SET*

    > Set the label the item will have.

  * Placeholder ( string ) *GET/SET*

    > On supported browsers show this value if no value is set.

  * iClass ( string ) *GET/SET*

    > Add additional this additional class to the enclosing div.

  * ValidateRequired ( bool ) *SET*

    > To be valid the submitted form must contain a non empty value for this item.

  * ValidateRegex ( regex_pattern ) *SET*

    > To be valid the submitted form must contain a non empty value that matches the specified regex pattern.

  * ValidateContentType ( bool ) *SET*

    > For file items the uploaded file must be one of the content types specified in the "accept" property

FileItem
--------

This class is not meant to be directly created by the user, it is returned by the other form classes when accessing the value of a file element.

  * FileName ( string ) *GET/SET*

    > Get or set the file name of the uploaded file.

  * ContentType () *GET*

    > Gets the MIME type of the uploaded file if it was provided by the browser.

  * Size () *GET*

    > Gets the size of the uploaded file in bytes.

  * Save( path )

    > Saves the file to disk at the path specified with the current filename.

  * Blob () *GET*

    > Returns the binary blob of the file so you could, for instance, store it in a database.