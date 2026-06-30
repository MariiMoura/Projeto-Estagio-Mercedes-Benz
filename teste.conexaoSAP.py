import win32com.client

sap = win32com.client.GetObject("SAPGUI")
app = sap.GetScriptingEngine
session = app.Children(0).Children(0)

print("Conectado com sucesso!")