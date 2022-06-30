' Local date and time formatter

Dim dtmNow
dtmNow = Now
' output format: yyyymmddhhnnss
wscript.echo (((year(dtmNow)*100 + month(dtmNow))*100 + day(dtmNow))*10000 + hour(dtmNow)*100 + minute(dtmNow))*100 + second(dtmNow)
