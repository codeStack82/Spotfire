import clr
clr.AddReference("Microsoft.Office.Interop.Outlook")
from System.Runtime.InteropServices import Marshal

mail= Marshal.GetActiveObject("Outlook.Application").CreateItem(0)
mail.Recipients.Add("email@gmail.com")
mail.Subject = "Subject here"
mail.Body = "Body here"
mail.Send();
