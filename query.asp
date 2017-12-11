<%
FUNCTION Query(sql)
   SET connection = Server.CreateObject("ADODB.Connection")
   SET cmd = Server.CreateObject("ADODB.Command")

   connection.ConnectionString = dbCon
   connection.Open()
   cmd.ActiveConnection = connection

   sql = AntiInject(sql)
   cmd.CommandText = sql
   SET Query = cmd.Execute()
END FUNCTION

FUNCTION AntiInject(LCCadVerificar)  
	Dim LCCadenaAntiInject
	LCCadenaAntiInject = ""
	
	Dim LNContador
	For LNContador = 1 To LEN(LCCadVerificar)
		If Mid(LCCadVerificar, LNContador, 1) = ";" Then
			LCCadenaAntiInject = Mid(LCCadVerificar, 1, LNContador - 1)
			
			AntiInject = LCCadenaAntiInject
			EXIT FUNCTION
		End If
	NEXT
	
	AntiInject = LCCadVerificar
END FUNCTION

FUNCTION ToJSON(qres)
   Dim json
   json = "["
   DO WHILE NOT qres.EOF
      json = json & "{"
      FOR i=0 TO qres.Fields.Count-1 STEP 1
         json = json & """" & qres.Fields(i).Name & """: """ & Trim(qres.Fields(i)) & """"
         IF i <> qres.Fields.Count-1 THEN
            json = json & ","
         END IF
      NEXT
      json = json & "}"
      qres.MoveNext()
      IF NOT qres.EOF THEN
         json = json & ","
      END IF
   LOOP
   json = json & "]"
   qres.Close()
   toJson = json
END FUNCTION

FUNCTION ToArray(qres, key)
   REDIM PRESERVE arr(1)
   DIM i,size
   i = 0
   size = 1
   DO WHILE NOT qres.EOF
      IF i > size THEN
         REDIM PRESERVE arr(size + 1)
         size = size + 1
      END IF
      arr(i) = qres(key)
      i = i + 1
      qres.MoveNext
   LOOP
   ToArray = arr
END FUNCTION
%>
