set webbrowser = createobject("internetexplorer.application")
webbrowser.statusbar = false
webbrowser.menubar = false
webbrowser.toolbar = false
webbrowser.visible = true

webbrowser.navigate("https://accounts.google.com/ServiceLogin/identifier?service=mail&passive=true&rm=false&continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&ss=1&scc=1&ltmpl=default&ltmplcache=2&emr=1&osid=1&flowName=GlifWebSignIn&flowEntry=AddSession")

wscript.sleep(5000)

webbrowser.document.all.item("identifierId").value = "tharindujaye69@gmail.com"
webbrowser.document.all.item("identifierNext").click 

webbrowser.document.all.item("password").value = "tharindu69"

wscript.sleep(100)

webbrowser.document.all.item("passwordNext").click 
