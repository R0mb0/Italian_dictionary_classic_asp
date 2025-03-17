<%
    Class italian_dictionary
        
        Dim terminator '-> terminatore character of the string
        Dim base_array '-> the array where save the infos
        Dim array_index '-> array index of last searched element
        Dim flag2 '-> variable used to some iterative functions 
        Dim absolute_index '-> index used on the load function for the position in the string 
        Dim temp_text '->variable where store text
        
        ' Initialization and destruction'
        Sub class_initialize()
            terminator = "-"
            base_array = Array()
            flag2 = false 
            absolute_index = 0
            temp_text = ""
        End Sub

        Sub class_terminate()
            terminator = Nothing 
            base_array = Nothing
            flag2 = Nothing
            absolute_index = Nothing
            temp_text = Nothing
            array_index = Nothing
        End Sub

        'Function to initialize the class and load the words 
        Public Function initialize(ByVal lists)
        Dim temp_array
        temp_array = Split(lists, " ")
        Dim temp 
        For Each temp in temp_array
            Select Case temp 
                Case "1000_parole_italiane_comuni" 
                    load_file("1000_parole_italiane_comuni.txt")
                Case "110000_parole_italiane_con_nomi_propri"
                    load_file("110000_parole_italiane_con_nomi_propri.txt")
                Case "280000_parole_italiane"
                    load_file("280000_parole_italiane.txt")
                Case "400_parole_composte"
                    load_file("400_parole_composte.txt")
                Case "60000_parole_italiane"
                    load_file("60000_parole_italiane.txt")
                Case "660000_parole_italiane"
                    load_file("660000_parole_italiane.txt")
                Case "9000_nomi_propri"
                    load_file("9000_nomi_propri.txt")
                Case "95000_parole_italiane_con_nomi_propri"
                    load_file("95000_parole_italiane_con_nomi_propri.txt")
                Case "coniugazione_verbi"
                    load_file("coniugazione_verbi.txt")
                Case "lista_38000_cognomi"
                    load_file("lista_38000_cognomi.txt")
                Case "lista_badwords"
                    load_file("lista_badwords.txt")
                Case "lista_cognomi"
                    load_file("lista_cognomi.txt")
                Case "parole_uniche"
                    load_file("parole_uniche.txt")
                Case Else 
                    Call Err.Raise(vbObjectError + 10, "italian_dictionary.class","initialize - This file: " & temp & ".txt, has not been funded")
            End Select 
        Next 
        End Function

        'Function to print a debug message 
        Private Function dp(message)
            Response.write "<br><h3> Debug print: " & message & " </h3><br>"
        End Function 

        'Function to convert a string into an array
        Private Function string_to_array(ByVal text)
            Dim length
            length = Len(text)
            Dim outArray() 
            Dim index 
            For index = 0 to length - 1
                Redim preserve outArray(length)
                outArray(index) = Left(Right(text,(length - index)), (1))
            Next 
            Redim preserve outArray(length - 1)
            string_to_array = outArray
        End Function

        'Ad an element in the array head
        Private Function add_base_element(ByVal element, ByRef array)
            Dim temp 
            temp =  UBound(array) + 1
            Redim Preserve array(temp)
            array(temp) = element
        End Function 

        'Return true (and save the index) if find the element, else false 
        Private Function search_base_element(ByVal element, ByRef array)
            Dim temp 
            Dim my_index 
            my_index = 0
            For Each temp In array 
                If IsArray(temp) Then 
                    If temp(0) = element Then 
                        array_index = my_index
                        search_base_element = true
                        Exit Function 
                    End If 
                End If 
                my_index = my_index + 1
            Next 
            search_base_element = false
        End Function 

        'This function ad an element 
        Private Function node(ByVal value)
            Dim temp_array(0)
            temp_array(0) = value
            node = temp_array
        End Function 

        'Function to add a word inside the array
        Private Function adding_word(ByVal word, ByVal index, ByRef array)
            If Not(index <= UBound(word)) Then 'If index is not valid 
                Exit Function
            End If
            If UBound(array) = "-1" Then 'If the array is empty
                Redim preserve array(0)
                array(0) = node(word(index))
                adding_word word, index + 1, array(0)
                Exit Function
            End If 
            If search_base_element(word(index), array) Then 
                adding_word word, index + 1, array(array_index)
                Exit Function
            Else
                add_base_element node(word(index)) ,array 
                adding_word word, index + 1, array(UBound(array))
                Exit Function
            End If 
        End Function 

        'Private function to add word in more efficient way if the function "add_text" has been invoked
        Private Function private_add_word(ByVal word)
            Dim my_word 
            my_word = word
            'If is case sentive 
            If Not case_sensitive Then 
                my_word = LCase(my_word)
            End If 
            my_word = string_to_array(my_word)
            add_base_element terminator, my_word  
            adding_word my_word, 0, base_array 
        End Function

        'Public function to add the words of a text 
        Private Function add_text(ByVal text)
            Dim temp 
            For Each temp In Split(text, " ")
                private_add_word(temp)
            Next
        End Function

        'Function to load a dictionary 
        Private Function load_file(file)
             'Read part 
            Dim fs
            Dim t
            Dim s
            Set fs = Server.CreateObject("Scripting.FileSystemObject")
            set t = fs.OpenTextFile(file, 1, false)
            s = t.ReadAll
            t.close
            Set t = Nothing
            Set fs = Nothing
            'Add dictionary
            add_text(s)
        End Function 

         'Function to retrieve the words inside the base array 
        Private Function write_array(ByRef array, ByVal flag1, ByVal chara)
            If flag1 then 'If I've found a character
                Dim my_flag
                my_flag = flag1 
                Dim temp 
                For Each temp In array 'explore the arrays 
                    If Not(my_flag) Then  
                        write_array = write_array & write_array(temp, my_flag, chara)
                    Else 
                        my_flag = Not(my_flag)
                    End If 
                Next 
            Else 
                If Not(IsArray(array(0))) Then 'if the first element is a character
                    If array(0) = terminator Then 'if I've found the terminator character
                        flag2 = true   
                        write_array = "; "
                    Else 
                        If flag2 Then 'If I've found a new word
                            flag2 = false
                            write_array = chara & array(0) & write_array(array, true, chara & array(0))
                        Else 
                            write_array = array(0) & write_array(array, true, chara & array(0))
                        End If 
                    End If 
                Else 'If the array is the base_array
                    For Each temp In array
                        write_array = write_array & write_array(temp, false, chara) & "<br>"
                    Next
                End If 
            End If 
        End Function 

        'Function to print all the elements inside the search tree 
        Public Function write_all_words()
            Response.write write_array(base_array, false, "")
        End Function 

        'Private function to find a word inside the tree
        Private Function find_word(ByVal word, ByVal index, ByRef array)
            If search_base_element(word(index), array) Then 
                If word(index) = terminator Then 
                    find_word = true
                    Exit Function 
                Else 
                    find_word = find_word(word, index + 1, array(array_index))
                    Exit Function
                End IF 
            End If 
            find_word = false 
        End Function 

        'Private Function to check if a word is in the memory
        Private Function private_is_present(ByVal word)
            Dim my_word
            my_word = word
            'If is case sentive 
            If Not case_sensitive Then 
                my_word = LCase(my_word)
            End If 
            my_word = string_to_array(word)
            add_base_element terminator, my_word 
            Dim temp 
            Dim index
            For Each temp In base_array
                If temp(0) = my_word(0) Then 
                private_is_present = find_word(my_word, 1, base_array(index))
                Exit Function 
                End If 
                index = index + 1
            Next
            private_is_present = false 
        End Function 

        'Function to check if a word is in the tree
        Public Function is_present(ByVal word)
            is_present = private_is_present(word)
        End Function 


        'Function to retrieve the rest of the words from the passed word -> added for perfromance reasons 
        Private Function add_characters(ByRef array, ByVal flag1, ByVal chara)
            If flag1 then 'If I've found a character
                Dim my_flag
                my_flag = flag1 
                Dim temp 
                For Each temp In array 'explore the arrays 
                    If Not(my_flag) Then  
                        add_characters = add_characters & add_characters(temp, my_flag, chara)
                    Else 
                        my_flag = Not(my_flag)
                    End If 
                Next 
            Else 
                If Not(IsArray(array(0))) Then 'if the first element is a character
                    If array(0) = terminator Then 'if I've found the terminator character
                        flag2 = true   
                        add_characters = " "
                    Else 
                        If flag2 Then 'If I've found a new word
                            flag2 = false
                            add_characters = chara & array(0) & add_characters(array, true, chara & array(0))
                        Else 
                            add_characters = array(0) & add_characters(array, true, chara & array(0))
                        End If 
                    End If 
                Else 'If the array is the base_array
                    For Each temp In array
                        add_characters = add_characters & add_characters(temp, false, chara)
                    Next
                End If 
            End If 
        End Function 

        'Function to print the rest of word passed to search_word
        Private Function retrieve_words(ByVal word, ByVal index, ByRef array, ByVal flag, ByVal chara)
            Dim temp 
            'If the word's character are spent
            If flag Then 
                retrieve_words = add_characters(array, false, chara)
                Exit Function
            Else
                If search_base_element(word(index), array) Then 
                    If index = UBound(word) Then 
                        retrieve_words = retrieve_words(word, index + 1, array(array_index), true, chara)
                        Exit Function 
                    End If 
                    retrieve_words = retrieve_words(word, index + 1, array(array_index), false, chara & word(index))
                    Exit Function 
                End If 
            End If 
        End Function 

        'Function to search a word inside the memory
        Public Function search_word(ByVal word, ByVal is_array)
            'In case of null argument 
            If word = " " and (Len(word) = 0) Then 
                search_word = " "
                Exit Function 
            End If 
            'If the word is present then exit 
            If private_is_present(word) Then 
                search_word = " "
                Exit Function
            End If 
            Dim my_word
            my_word = word
            'If is case sentive 
            If Not case_sensitive Then 
                my_word = LCase(my_word)
            End If 
            my_word = string_to_array(word)
            If is_array Then 
                search_word = Split(retrieve_words(my_word, 0, base_array, false, ""), " ")
            Else 
                search_word = Replace(retrieve_words(my_word, 0, base_array, false, ""), " ", "; ")
            End If 
        End Function 

        'Private Funtion to serialize an array 
        Private Function serialize_array(ByVal array)
            Dim my_string 
            If IsArray(array) Then 
                my_string = "["
                Dim temp 
                For Each temp in array 
                    my_string = my_string & serialize_array(temp)
                Next 
                my_string = my_string & "]"
                serialize_array = my_string
            Else 
                serialize_array = array
            End If 
        End Function 

        'Function to save the tree in a file 
        Public Function save_tree(ByVal path)
            Dim fso 
            Set fso = Server.CreateObject("Scripting.FileSystemObject")
            If Not(fso.FileExists(path) Or fso.FolderExists(path)) Then 
                Call Err.Raise(vbObjectError + 10, "italian_dictionary.class","load_tree - The path is not valid")
            End If 
            Set fso = Nothing
            Dim temp_string
            temp_string = serialize_array(base_array)
            'Save string to file 
            Dim fs
            Dim f
            Set fs = Server.CreateObject("Scripting.FileSystemObject")
            Set f = fs.CreateTextFile(path, true)
            f.write(temp_string)
            f.close
            Set f = Nothing
            Set fs = Nothing
            'End saving string to file 
            'Return for debug purpose
            save_tree = temp_string
        End Function 

       'Function to get a character from a string using a index.
        Private Function get_character(ByVal index, ByRef text)
            If IsNumeric(index) and index <= Len(text) Then 
                get_character = Left(Right(text,(Len(text) - index)), (1))
            Else 
                Call Err.Raise(vbObjectError + 10, "italian_dictionary.class","get_character - Irregular Index: " & text)
            End If 
        End Function 

        'Function to add arrays following the text
        Private Function add_from_text()
            Dim returnArray() '<---- To avoid this error: This array is fixed or temporarily locked
            ReDim returnArray(0)
            returnArray(0) = Null 
            Dim character
            character = get_character(absolute_index, temp_text)
            Do While character <> "]"
                If character = "[" Then
                    absolute_index = absolute_index + 1 '<---- Had read
                    If IsNull(returnArray(0)) Then 
                        returnArray(0) = add_from_text()
                    Else 
                        ReDim Preserve returnArray(UBound(returnArray) + 1)
                        returnArray(UBound(returnArray)) = add_from_text()
                    End If 
                Else 
                    absolute_index = absolute_index + 1 '<---- Had read
                    If IsNull(returnArray(0)) Then 
                        returnArray(0) = character
                    Else 
                        ReDim Preserve returnArray(UBound(returnArray) + 1)
                        returnArray(UBound(returnArray)) = character
                    End If 
                End If 
                character = get_character(absolute_index, temp_text)
            Loop 
            absolute_index = absolute_index + 1 '<---- Had read
            add_from_text = returnArray
        End Function 

        'Funtion to load the tree from a file 
        Public Function load_tree(ByVal path)
            Dim fso 
            Set fso = Server.CreateObject("Scripting.FileSystemObject")
            If Not(fso.FileExists(path) Or fso.FolderExists(path)) Then 
                Call Err.Raise(vbObjectError + 10, "italian_dictionary.class","load_tree - The path is not valid")
            End If 
            Set fso = Nothing
            'Read part 
            Dim fs
            Dim t
            Dim s
            Set fs = Server.CreateObject("Scripting.FileSystemObject")
            set t = fs.OpenTextFile(path, 1, false)
            s = t.ReadAll
            t.close
            Set t = Nothing
            Set fs = Nothing
            'End read part
            'If necessary format base array
            If UBound(base_array) > 0 Then 
                Redim base_array(0)
                base_array(0) = null 
            End If 
            Dim temp_string 
            temp_string = Left(Right(s, Len(s)- 1), Len(s)- 2) '<------ Removed the first and the last parenthesis
            Dim length
            length = Len(temp_string)
            temp_text = temp_string
            'Starting Loop 
            Do While absolute_index < length
                If get_character(absolute_index, temp_text) = "[" Then
                    absolute_index = absolute_index + 1 '<---- Had read
                    If IsNull(base_array(0)) Then '<---- If the array has been just initialized 
                        base_array(0) = add_from_text()
                    Else 
                        Redim Preserve base_array(UBound(base_array) + 1)
                        base_array(UBound(base_array)) = add_from_text()
                    End If 
                Else
                    'If the string is not regular
                    Call Err.Raise(vbObjectError + 10, "italian_dictionary.class","load_tree - Irregular File: " & get_character(absolute_index, temp_string))
                End If 
            Loop 
        End Function
    End Class 
%> 