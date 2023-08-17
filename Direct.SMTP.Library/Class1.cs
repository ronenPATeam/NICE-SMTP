using System;
using log4net;
using Direct.Shared;
using System.Net.Mail;


namespace Direct.SMTP.Library
{
    [DirectSealed]
    [DirectDom("SMTP")]
    [ParameterType(false)]
    public class SMTPLibrary
    {

        private static readonly ILog logArchitect = LogManager.GetLogger(Loggers.LibraryObjects);

        //Sending Email messages using SMTP server 
        [DirectDom("Send an email using SMTP")]
        [DirectDomMethod("Send {Message}")]
        [MethodDescriptionAttribute("Send SMTP Email Message using Email message Type")]
        public static string sendEmail(EmailMessage msg)
        {
            if (logArchitect.IsDebugEnabled)
            {
                logArchitect.DebugFormat("SMTP Library - attempting to send message with subject {0}", msg.Subject);


            }

            //Smtp client properties 
            SmtpClient client = new SmtpClient();
            client.Host = msg.ServerDetails;
            client.Port = msg.Port;
            client.EnableSsl = msg.ESSL;
            client.Timeout = msg.sendTimeOut;
            client.DeliveryMethod = SmtpDeliveryMethod.Network;
            if (msg.UseDefaultCredentials)
            {
                if (logArchitect.IsDebugEnabled)
                {
                    logArchitect.DebugFormat("SMTP Library - Using deafult credentials");


                }
                client.UseDefaultCredentials = true;
            }
            else
            {
                client.UseDefaultCredentials = false;
                client.Credentials = new System.Net.NetworkCredential(msg.UserName, msg.Password);
            }

            //Email Message Properties
            using (MailMessage mm = new MailMessage())
            {
                mm.BodyEncoding = System.Text.UTF8Encoding.UTF8;
                mm.Sender = new System.Net.Mail.MailAddress(msg.Sender);
                mm.Body = msg.Body;
                mm.From = new System.Net.Mail.MailAddress(msg.Sender);
                mm.IsBodyHtml = msg.isHtml;
                mm.Subject = msg.Subject;
                mm.DeliveryNotificationOptions = DeliveryNotificationOptions.OnFailure;

                if (msg.isHtml && msg.isSignature)
                {
                    string htmlBodyWithSignature = msg.Body;
                    AlternateView htmlView = AlternateView.CreateAlternateViewFromString(htmlBodyWithSignature, null, "text/html");

                    LinkedResource signatureImage = new LinkedResource(msg.signaturePath, "image/png");
                    signatureImage.ContentId = "signature_image";
                    htmlView.LinkedResources.Add(signatureImage);

                    mm.AlternateViews.Add(htmlView);
                }

                //Adding Recipients and CC list
                MailAddressCollection tos = new MailAddressCollection();
                foreach (string addr in msg.Recipients)
                {
                    if (logArchitect.IsDebugEnabled)
                    {
                        logArchitect.DebugFormat("SMTP Library - attempting to add address {0}", addr);
                    }
                    mm.To.Add(addr);
                    logArchitect.DebugFormat("SMTP Library - attempting to add address {0}", addr);
                }
                MailAddressCollection tosCC = new MailAddressCollection();
                foreach (string addr in msg.RecipientsCC)
                {
                    if (logArchitect.IsDebugEnabled)
                    {
                        logArchitect.DebugFormat("SMTP Library - attempting to add addressCC {0}", addr);
                    }
                    mm.CC.Add(addr);
                    logArchitect.DebugFormat("SMTP Library - attempting to add addressCC {0}", addr);
                }

                //Adding attachemets
                foreach (string att in msg.Attachements)
                {
                    if (logArchitect.IsDebugEnabled)
                    {
                        logArchitect.DebugFormat("SMTP Library - attempting to add attachement {0}", att);
                    }
                    mm.Attachments.Add(new Attachment(att));
                    logArchitect.DebugFormat("SMTP Library - attempting to add attachments {0}", att);
                }

                //Sending the Email Message and returning errors or success result
                try
                {
                    client.Send(mm);
                    client.Dispose();
                }
                catch (Exception e)
                {
                    logArchitect.Error("SMTP Library - Send SMTP email error. ", e);
                    client.Dispose();
                    return e.ToString();
                }
                return "Message Sent Successfully";
            }

        }



    }

    [DirectDom("Email Message")]
    public class EmailMessage : DirectComponentBase
    {

        protected PropertyHolder<string> _Subject = new PropertyHolder<string>("Subject");
        protected PropertyHolder<string> _ServerDetails = new PropertyHolder<string>("ServerDetails");
        protected CollectionPropertyHolder<string> _Recipients = new CollectionPropertyHolder<string>("Recipients");
        protected CollectionPropertyHolder<string> _RecipientsCC = new CollectionPropertyHolder<string>("RecipientsCC");
        protected CollectionPropertyHolder<string> _Attachements = new CollectionPropertyHolder<string>("Attachements");
        protected PropertyHolder<string> _Body = new PropertyHolder<string>("Body");
        protected PropertyHolder<string> _UserName = new PropertyHolder<string>("UserName");
        protected PropertyHolder<string> _Password = new PropertyHolder<string>("Password");
        protected PropertyHolder<string> _Sender = new PropertyHolder<string>("Sender");
        protected PropertyHolder<int> _Port = new PropertyHolder<int>("Port");
        protected PropertyHolder<bool> _ESSL = new PropertyHolder<bool>("ESSL");
        protected PropertyHolder<bool> _isHtml = new PropertyHolder<bool>("isHtml");
        protected PropertyHolder<bool> _useDefaultCredentials = new PropertyHolder<bool>("UseDefaultCredentials");
        protected PropertyHolder<int> _sendTimeOut = new PropertyHolder<int>("sendTimeOut");
        protected PropertyHolder<bool> _isSignature = new PropertyHolder<bool>("isSignature");
        protected PropertyHolder<string> _signaturePath = new PropertyHolder<string>("signaturePath");

        //List property holder
        [DirectDomAttribute("Recipients")]
        [DesignTimeInfo("Recipients")]
        public DirectCollection<string> Recipients
        {
            get
            {

                return this._Recipients.TypedValue;

            }
            set
            {
                this._Recipients.TypedValue = value;
            }
        }

        [DirectDomAttribute("RecipientsCC")]
        [DesignTimeInfo("RecipientsCC")]
        public DirectCollection<string> RecipientsCC
        {
            get
            {

                return this._RecipientsCC.TypedValue;

            }
            set
            {
                this._RecipientsCC.TypedValue = value;
            }
        }

        [DirectDomAttribute("Attachements")]
        [DesignTimeInfo("Attachements")]
        public DirectCollection<string> Attachements
        {
            get
            {

                return this._Attachements.TypedValue;

            }
            set
            {
                this._Attachements.TypedValue = value;
            }
        }

        //Simple properties
        [DirectDom("Subject")]
        public string Subject
        {
            get { return _Subject.TypedValue; }
            set { this._Subject.TypedValue = value; }
        }


        [DirectDom("ServerDetails")]
        public string ServerDetails
        {
            get { return _ServerDetails.TypedValue; }
            set { this._ServerDetails.TypedValue = value; }
        }
        [DirectDom("Body")]
        public string Body
        {
            get { return _Body.TypedValue; }
            set { this._Body.TypedValue = value; }
        }

        [DirectDom("UserName")]
        public string UserName
        {
            get { return _UserName.TypedValue; }
            set { this._UserName.TypedValue = value; }
        }
        [DirectDom("Password")]
        public string Password
        {
            get { return _Password.TypedValue; }
            set { this._Password.TypedValue = value; }
        }

        [DirectDom("Sender")]
        public string Sender
        {
            get { return _Sender.TypedValue; }
            set { this._Sender.TypedValue = value; }
        }
        [DirectDom("Port")]
        public int Port
        {
            get { return _Port.TypedValue; }
            set { this._Port.TypedValue = value; }
        }
        [DirectDom("ESSL")]
        public bool ESSL
        {
            get { return _ESSL.TypedValue; }
            set { this._ESSL.TypedValue = value; }
        }
        [DirectDom("isHtml")]
        public bool isHtml
        {
            get { return _isHtml.TypedValue; }
            set { this._isHtml.TypedValue = value; }
        }
        [DirectDom("Use Default Credentials")]
        public bool UseDefaultCredentials
        {
            get { return _useDefaultCredentials.TypedValue; }
            set { this._useDefaultCredentials.TypedValue = value; }
        }
        [DirectDom("sendTimeOut")]
        public int sendTimeOut
        {
            get { return _sendTimeOut.TypedValue; }
            set { this._sendTimeOut.TypedValue = value; }
        }

        [DirectDom("isSignature")]
        public bool isSignature
        {
            get { return _isSignature.TypedValue; }
            set { this._isSignature.TypedValue = value; }
        }

        [DirectDom("signaturePath")]
        public string signaturePath
        {
            get { return _signaturePath.TypedValue; }
            set { this._signaturePath.TypedValue = value; }
        }


        //constructors
        private void initLists()
        {
            //init lists
            this.Recipients = new DirectCollection<string>(this.Project);
            this.RecipientsCC = new DirectCollection<string>(this.Project);
            this.Attachements = new DirectCollection<string>(this.Project);
        }
        public EmailMessage()
        {
            initLists();
        }
        public EmailMessage(Direct.Shared.IProject project)
            : base(project)
        {
            initLists();
        }
    }
}
