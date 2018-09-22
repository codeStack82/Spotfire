import clr
clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import MessageBox, MessageBoxButtons
from System.Windows.Forms import DialogResult
import Spotfire.Dxp.Application.Visuals.VisualContent as visualcontent
from Spotfire.Dxp.Application import Page
from Spotfire.Dxp.Application.Visuals import VisualContent
from System import IO, Net, Environment, String, Array
from System.Net import Mail, Mime
from System.Drawing import Bitmap, Graphics, Rectangle, Point
from System.IO import Path, File, Stream, MemoryStream, StreamReader, StreamWriter
from System.Text import Encoding
from Spotfire.Dxp.Application.Filters import CheckBoxFilter 
from System import IO, Net, DateTime
from System.Net import Mail, Mime
from System.Text import Encoding


# Info **********************************
# Last updated:		9/22/18 
# Created by: 	  Tyler Hunt

# TODO: Need to get var dynamically from export table
wellName = "well name"
rigName = "rig numver"

# TODO: Build out dynamic file addition to the email

def sendEmail():
	
	# Configurations
	SMTPClient = "mail.example.com"
	SMTPPort = 25
	useEncription = False
	useCredentials = True
	userEmail = Environment.UserName + "@example.com"

	# Prepare email
	MyMailMessage = Mail.MailMessage()
	MyMailMessage.From = Mail.MailAddress(userEmail)

	#send email to self
	#MyMailMessage.To.Add(userEmail) 

	#TODO: build list for multiple to addresses
	MyMailMessage.To.Add("your.name@ex.com")

	MyMailMessage.Subject = "Subject line " + wellName  + " on " + rigName 
	MyMailMessage.IsBodyHtml = True
	
	# Create HTML body of email
	htmlBody = "<html><body>"
	htmlBody += "<span><p>All, <br><br> </p></span>"
	htmlBody += "<span><p>Attached is the example " + wellName  + " on " + rigName +  "<br><br> </p></span>"
	
	#for image in imageList:
		#htmlBody += "<span><p><b>" + image[0] + "</b></p><img src=\"cid:" + image[3] + "\" /></span>"

	htmlBody += "</body></html>"
	#print htmlBody
	
	# Create alternate view for html email
	altViewHTML = Mail.AlternateView.CreateAlternateViewFromString(htmlBody, None, Mime.MediaTypeNames.Text.Html)
	
	# Add the alternate views instead of using the 'MailMessage.Body'
	MyMailMessage.AlternateViews.Add(altViewHTML)
	
	# Create SMTPClient
	SMTPServer = Mail.SmtpClient(SMTPClient)
	SMTPServer.Port = SMTPPort
	
	# Send email
	if useCredentials:
		SMTPServer.UseDefaultCredentials = True
	SMTPServer.EnableSsl = useEncription
	SMTPServer.Send(MyMailMessage)
	print "The email has been sent"


# Main*****************************************
# Message box code
dialogResult = MessageBox.Show("Are you sure you want to email the OSC Roadmap for the below well? \n\n\tWell:\t" +  wellName  + " on \n\t  Rig:\t" + rigName, "OSC Roadmap Confirmation", MessageBoxButtons.YesNo)

if dialogResult == DialogResult.Yes:
	sendEmail() 
	MessageBox.Show("The OSC Roadmap for the " + wellName + " has been sent to " + rigName)

else:
    MessageBox.Show("There has been some sort of error. Please try again never!")
