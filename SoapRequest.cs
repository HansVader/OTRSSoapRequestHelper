using System;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Security.Authentication;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

namespace OtrsSoapRequest
{
    /// <summary>
    /// This class is a helper class for making soap request to the OTRS systems.
    /// SOAP standard 1.2 is used for all calls. No encryption is implementet yet.
    /// </summary>
    public class SoapRequestHelper
    {
        /// <summary>
        /// Gets or sets the namespace for the SOAP messages
        /// </summary>
        public XNamespace Namespace { get; set; }

        public SoapRequestHelper()
        {
            // If set to true client requests will add a Expect100Continue header. 
            // We don't want this behaviour since we are usually sending one small request
            // and the server should process this immediately.
            // See here for detailed information: 
            // https://msdn.microsoft.com/de-de/library/system.net.servicepointmanager.expect100continue(v=vs.110).aspx#Anchor_1
            System.Net.ServicePointManager.Expect100Continue = false;
            Namespace = "http://www.otrs.org/TicketConnector/";
        }

        #region SessionCreate
        /// <summary>
        /// Gets the XML template for the SessionCreate message.
        /// </summary>
        /// <returns>Template for SessionCreate message</returns>
        private XDocument GetSoapSessionCreateMessage()
        {
            var assembly = Assembly.GetExecutingAssembly();
            XDocument soapMessage = XDocument.Load(assembly.GetManifestResourceStream("OTRSSoapRequestHelper.XML.SessionCreate.xml"));
            return soapMessage;
        }


        /// <summary>
        /// Fills the leafs of a given XML template with values from the parameters. 
        /// </summary>
        /// <param name="soapMessage">XML document template to be filled</param>
        /// <param name="userID">UserID of the OTRS agent</param>
        /// <param name="password">Password of the OTRS agent</param>
        /// <returns>Pepared XML document with filled leaf values</returns>
        private XDocument PrepareSessionCreateSoapMessage(XDocument soapMessage, string userID, string password)
        {
            var element = soapMessage.Descendants(Namespace + "UserLogin").FirstOrDefault();
            if (element != null)
                element.Value = userID;

            element = soapMessage.Descendants(Namespace + "Password").FirstOrDefault();
            if (element != null)
                element.Value = password;

            return soapMessage;
        }

        /// <summary>
        /// Creates an OTRS session id which will be used for future webservice calls like <code>TicketUpdate</code>
        /// </summary>
        /// <param name="userID"></param>
        /// <param name="password"></param>
        /// <returns>SessionID which can be used for future calls</returns>
        /// <exception cref="AuthenticationException">Username or password is invalid</exception>
        public async Task<string> CreateOtrsSession(string userID, string password, string hostName)
        {
            if (String.IsNullOrEmpty(userID))
                throw new ArgumentException("Benutzer darf nicht leeer sein");

            if (String.IsNullOrEmpty(password))
                throw new ArgumentException("Passwort darf nicht leer sein!");

            if (String.IsNullOrEmpty(hostName))
                throw new ArgumentException("Hostname darf nicht leer sein!");

            var soapEnvelopeXml = GetSoapSessionCreateMessage();
            soapEnvelopeXml = PrepareSessionCreateSoapMessage(soapEnvelopeXml, userID, password);
            var soapRequest = CreateSoapRequest("http://" + hostName + "/otrs/nph-genericinterface.pl/Webservice/GenericTicketConnector", "SessionCreate");
            await InsertSoapEnvelopeIntoSoapRequest(soapEnvelopeXml, soapRequest);

            using (var stringWriter = new StringWriter())
            {
                using (var xmlWriter = XmlWriter.Create(stringWriter))
                {
                    soapEnvelopeXml.WriteTo(xmlWriter);
                }
            }

            // begin async call to web request.
            var asyncResult = await soapRequest.GetResponseAsync();

            // get the response from the completed web request.
            var responseStream = asyncResult.GetResponseStream();

            using (var reader = new StreamReader(responseStream))
            {
                XDocument soapResultDoc = XDocument.Parse(reader.ReadToEnd());
                if (!ContainsError(soapResultDoc))
                {
                    return soapResultDoc.Descendants(Namespace + "SessionID").FirstOrDefault().Value.ToString();
                }
                else
                {
                    var errorMessage = soapResultDoc.Descendants(Namespace + "ErrorMessage").FirstOrDefault().Value;
                    var errorCode = soapResultDoc.Descendants(Namespace + "ErrorCode").FirstOrDefault().Value;
                    throw new Exception(errorCode + "\r" + errorMessage);
                }
            }
        }

        #endregion

        #region TicketUpdate
        private XDocument GetSoapTicketUpdateMessage()
        {
            var assembly = Assembly.GetExecutingAssembly();
            XDocument soapMessage = XDocument.Load(assembly.GetManifestResourceStream("SoapRequest.XML.TicketUpdate.xml"));
            return soapMessage;
        }

        private XDocument PrepareTicketUpdateSoapMessage(XDocument soapMessage, string sessionID, string ticketNumber, string message, double timeUnit)
        {
            var element = soapMessage.Descendants("SessionID").FirstOrDefault();
            if (element != null)
                element.Value = sessionID;

            element = soapMessage.Descendants("TicketNumber").FirstOrDefault();
            if (element != null)
                element.Value = ticketNumber;

            element = soapMessage.Descendants("Body").FirstOrDefault();
            if (element != null)
                element.Value = message;

            element = soapMessage.Descendants("TimeUnit").FirstOrDefault();
            if (element != null)
                element.Value = timeUnit.ToString();

            return soapMessage;
        }

        /// <summary>
        /// Appends a note to a given ticketnumber.
        /// </summary>
        /// <param name="sessionID">Will be used to authenticate a user</param>
        /// <param name="ticketNumber">Ticketnumber which should be updated</param>
        /// <param name="message">Message body of the note</param>
        /// <param name="timeUnit">Amount a agent spent for this task</param>
        /// <returns>ArticleID if successfull</returns>
        public async Task<long> UpdateOtrsTicket(string sessionID, string ticketNumber, string message, double timeUnit, string hostName)
        {
            if (String.IsNullOrEmpty(sessionID))
                throw new ArgumentException("sessionID darf nicht leer sein.\r" +
                                            "CreateOtrsSession() sollte zuerst aufgerufen werden um eine OTRS SessionID zu generieren.");

            if (String.IsNullOrEmpty(ticketNumber))
                throw new ArgumentException("Benutzer darf nicht leeer sein");

            if (String.IsNullOrEmpty(message))
                throw new ArgumentException("Passwort darf nicht leer sein!");

            if (String.IsNullOrEmpty(hostName))
                throw new ArgumentException("Der Hostname darf nicht leer sein!");


            var soapEnvelopeXml = GetSoapTicketUpdateMessage();
            soapEnvelopeXml = PrepareTicketUpdateSoapMessage(soapEnvelopeXml, sessionID, ticketNumber, message, timeUnit);
            var soapRequest = CreateSoapRequest("http://" + hostName + "/otrs/nph-genericinterface.pl/Webservice/GenericTicketConnector", "TicketUpdate");
            await InsertSoapEnvelopeIntoSoapRequest(soapEnvelopeXml, soapRequest);

            using (var stringWriter = new StringWriter())
            {
                using (var xmlWriter = XmlWriter.Create(stringWriter))
                {
                    soapEnvelopeXml.WriteTo(xmlWriter);
                }
            }

            // begin async call to web request.
            var asyncResult = await soapRequest.GetResponseAsync();

            // get the response from the completed web request.
            var responseStream = asyncResult.GetResponseStream();

            using (var reader = new StreamReader(responseStream))
            {
                XDocument soapResultDoc = XDocument.Parse(reader.ReadToEnd());
                if (!ContainsError(soapResultDoc))
                {
                    return Convert.ToInt64(soapResultDoc.Descendants(Namespace + "ArticleID").FirstOrDefault().Value);
                }
                else
                {
                    var errorMessage = soapResultDoc.Descendants(Namespace + "ErrorMessage").FirstOrDefault().Value;
                    var errorCode = soapResultDoc.Descendants(Namespace + "ErrorCode").FirstOrDefault().Value;
                    throw new Exception(errorCode + "\r" + errorMessage);
                }
            }
        }
        #endregion TicketUpdate

        #region Generel
        /// <summary>
        /// Creates a HTTP requests and sets the SOAPAction header.
        /// The use of a proxy is disabled.
        /// </summary>
        /// <param name="url">URL of the webservice endpoint</param>
        /// <param name="action">Name of the SOAP action</param>
        /// <returns></returns>
        private HttpWebRequest CreateSoapRequest(string url, string action)
        {
            var webRequest = WebRequest.CreateHttp(url);
            // Do not use a proxy for this.
            webRequest.Proxy = null;
            // Add the SOAPAction header with namespace. 
            // The seperation character can be configured on the remote system.
            webRequest.Headers.Add("SOAPAction", Namespace.ToString() + "#" + action);
            webRequest.ContentType = "application/soap+xml; charset=utf-8";
            webRequest.Method = "POST";
            return webRequest;
        }

        private async Task InsertSoapEnvelopeIntoSoapRequest(XDocument soapEnvelopeXml, HttpWebRequest webRequest)
        {
            try
            {
                using (Stream stream = await webRequest.GetRequestStreamAsync())
                {
                    soapEnvelopeXml.Save(stream);
                }
            }
            catch (Exception exep)
            {
                throw exep;
            }
        }

        /// <summary>
        /// Checks if the XML answer from the remote system contains any error leafs
        /// </summary>
        /// <param name="document">XML document from the remote system</param>
        /// <returns>True if the documents contains an error leafs. False if not.</returns>
        private bool ContainsError(XDocument document)
        {
            return document.Descendants(Namespace + "Error").Any();
        }
        #endregion
    }
}
