# Italian dictionary in Classic ASP

[![Codacy Badge](https://app.codacy.com/project/badge/Grade/83fb604e6e074fb0b7f33dada989aa73)](https://app.codacy.com/gh/R0mb0/Italian_dictionary_classic_asp/dashboard?utm_source=gh&utm_medium=referral&utm_content=&utm_campaign=Badge_grade)

[![Maintenance](https://img.shields.io/badge/Maintained%3F-yes-green.svg)](https://github.com/R0mb0/Italian_dictionary_classic_asp)
[![Open Source Love svg3](https://badges.frapsoft.com/os/v3/open-source.svg?v=103)](https://github.com/R0mb0/Italian_dictionary_classic_asp)
[![MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/license/mit)

[![Donate](https://img.shields.io/badge/PayPal-Donate%20to%20Author-blue.svg)](http://paypal.me/R0mb0)

## `italian_dictionary.class.asp`'s avaible functions

- Initialize the class with the necessary words -> `Public Function initialize(ByVal lists)`
  > where `list` must be a string with the list of dictionary files to add. The files inside the string must be separated with space.  
  > The dictionary files to choose are:
  > - 1000_parole_italiane_comuni
  > - 110000_parole_italiane_con_nomi_propri
  > - 280000_parole_italiane
  > - 400_parole_composte
  > - 60000_parole_italiane
  > - 660000_parole_italiane
  > - 9000_nomi_propri
  > - 95000_parole_italiane_con_nomi_propri
  > - coniugazione_verbi
  > - lista_38000_cognomi
  > - lista_badwords
  > - lista_cognomi
  > - parole_uniche
  >
  > **This files has been taken from [paroleitaliane](https://github.com/napolux/paroleitaliane)**
- Print all elements of dictionary -> `Public Function write_all_words()`
- Check if a word is in the dictionary -> `Public Function is_present(ByVal word)`
- Search a word inside the dictionary -> `Public Function search_word(ByVal word, ByVal is_array)`
  > **This function is usefull to search a word inside the dictionary, for example if `word` = "cas" the function will return: "casa", "casina" and "casona"**
  > - `word` is the word to search, it could be a part of a word, if an entire word is passed to the function, the function will return null.
  > - `is_array` change the output of the function, if `true` the function will return an array with all results, else, will be returned a string
- Function to save to file the dictionary state -> `Public Function save_dictionary(ByVal path)`
  > **Where `path` is the string with the location with the file to save location**
- Function to load the saved state dictionary in a file -> `Public Function load_dictionary(ByVal path)`
  > **Where `path` is the string with the location with the file to load location**
  
## How to use

> From: `Test.asp`

1. Initialize the class
   ```asp
   <%@LANGUAGE="VBSCRIPT"%>
   <!--#include file="italian_dictionary.class.asp"-->
   <% 
      Dim dictionary
      Set dictionary = new italian_dictionary
      dictionary.initialize("1000_parole_italiane_comuni 400_parole_composte")
   ```
2. Fill the dictionary
   - ```asp
     dictionary.initialize("1000_parole_italiane_comuni 400_parole_composte")
     ```
   - Or
     ```asp
     dictionary.load_dictionary("path")
     ```
2. > Save the state of dictionary
   ```asp
   dictionary.save_dictionary("path")
   ```
3. Interrogate the tree   
   Possibilities:
   - Check if a word is in the dictionary
     ```asp
     dictionary.is_present("casa")
     ```
   - Search a word inside the tree
     ```asp
       dictionary.search_word("cas", false)
     %>
     ``` 
