# ATTACHMENTS-to-an-email-in-Outlook-with-Python

The first step is to create an outlook instance and mail item. We'll use the example from before except with the addition of the pathlib library. I'll explain why in the next step.

import win32com.client as client
import pathlib

outlook = client.Dispatch('Outlook.Application')
message = outlook.CreateItem(0)

message.To = 'shubhamojha808@gmail.com'
message.CC = 'shubhamojha809@gmail.com'


message.Subject = 'Happy Birthday"'
message.Body = ""
Adding an attachment
When you add an attachment in Outlook, you must use the absolute path of the file. The absolute path represents the entire file path from root to file name. Often you can get by with using the relative path as long as the file is in the same directly as a python file, but not so with an Outlook attachment. pathlib contains some very handy methods for dealing with this situation... specifically, the absolute() method.

We are going to attach two files to this email.

cake_path = pathlib.Path('birthday-cake.jpg')
cert_path = pathlib.Path('certificate.jfif')
The next step is to get the absolute path of the file. Because the absolute method returns a WindowsPath object, we need to wrap the result in str to convert it to a string.

cake_absolute = str(cake_path.absolute())
cert_absolute = str(cert_path.absolute())
Finally, we can use the Add method of the Attachments property to add this file to the email.

# add the birthday cake file
image = message.Attachments.Add(cake_absolute)

# add the certificate file
message.Attachments.Add(cert_absolute)
Now both files are attached to the email

Embedding images
You may be wondering, why did we save a reference to the birthday cake image, but not the certificate image when we created the attachment? Great question! The reason is that in order to embed the image into the HTML body of the email, we need to reference the images CID (content ID). We'll need a reference to the image attachment in order to make some changes to it in just a moment.

Look at the adjusted HTML Body template below, and you'll notice that I've replaced the URL link used in the last exercise with a reference to the CID (content ID).

html_body = """
    <div>
        <h1 style="font-family: 'Lucida Handwriting'; font-size: 56; font-weight: bold; color: #9eac9c;"> Happy Birthday!! </h1>
        <span style="font-family: 'Lucida Sans'; font-size: 28; color: #8d395c;"> Wishing you all the best on your birthday!! </span>
    </div><br>
    <div>
        <img src="cid:cake-img" width=50%>
    </div>
    """
Outlook will automatically generate a content id for the image, so we need to access and change that property with the property accessor to match what we've called it in our html template. You can read more about this on Microsoft's website. In order to access the content ID, we need to reference the code associated with the content id. I've included it below so you do not need to look it up.

image.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "cake-img")
Now that the CID has been changed, we can now update the HTMLBody of the message.

message.HTMLBody = html_body
You are now ready to save or send your email, which has both an embedded image and an attached certificate.
