<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="italian_dictionary.class.asp"-->
<% 
   Dim dictionary
   Set dictionary = new italian_dictionary
   dictionary.initialize("1000_parole_italiane_comuni 400_parole_composte")

   Response.write dictionary.is_present("casa")
   Response.write dictionary.search_word("cas", false)

   dictionary.save_dictionary("dictionary")
%>
