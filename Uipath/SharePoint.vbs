' en las actividades que uipath para hacer alguna funcionalidad vs sharepoint
' es muy importante que se tenga un filtro, esto permitira depurar la informacion que trae
' a continuacion se muestra un ejemplo de ocmo se puede realizar el filtro

'the aplication filter one 
"fields/" + nameColumn+ " eq '" + valueFilter + "'"

'the aplication filter two
"fields/" + nameColumn+ " eq '" + valueFilter  + "' and " +
"fields/nameColumn eq '" + valueFilter  + "'"

'the aplication filter three
"fields/" + nameColumn+ " eq '" + valueFilter  + "' and " + 
"fields/nameColumneq '" + valueFilter  + "' and " + 
"fields/nameColumneq '" + valueFilter  + "'"