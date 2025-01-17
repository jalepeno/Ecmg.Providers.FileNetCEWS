using System;
using System.Collections.Generic;
using System.Globalization;
using System.ServiceModel;
using System.ServiceModel.Channels;
using System.ServiceModel.Security;
using System.Text;
using CPEServiceReference;
 
namespace Documents.Providers.FileNetCEWS
{

    public class WSIUtil
    {
        private static Localization localization;
        private static FNCEWS40PortTypeClient port;

        public static FNCEWS40PortTypeClient ConfigureBinding(String user, String password, String uri)
        {


            EndpointAddress endpoint = new EndpointAddress(uri);

            var bindingElementCollection = new BindingElementCollection();
            var securityBindingElement = SecurityBindingElement.CreateUserNameOverTransportBindingElement();
            securityBindingElement.IncludeTimestamp = false;
            bindingElementCollection.Add(securityBindingElement);

            var encoding = new MtomMessageEncoderBindingElement(new TextMessageEncodingBindingElement
            {
                MessageVersion = MessageVersion.CreateVersion(EnvelopeVersion.Soap12, AddressingVersion.None)
            });


            bindingElementCollection.Add(encoding);

            if (uri.ToLower().Contains("https"))
            {
              var transportBindingElement = new HttpsTransportBindingElement { MaxReceivedMessageSize = 2147483647, MaxBufferSize = 2147483647 };
              bindingElementCollection.Add(transportBindingElement);
            }
            else
            {
              var transportBindingElement = new HttpTransportBindingElement { MaxReceivedMessageSize = 2147483647, MaxBufferSize = 2147483647 };
              bindingElementCollection.Add(transportBindingElement);
            }

            var customBinding = new CustomBinding(bindingElementCollection) { ReceiveTimeout = new TimeSpan(TimeSpan.TicksPerDay) };// 100 nanonsecond units, make it 1 day

            customBinding.SendTimeout = customBinding.ReceiveTimeout;


            port = new FNCEWS40PortTypeClient(customBinding, endpoint);

            port.ClientCredentials.UserName.UserName = user;
            port.ClientCredentials.UserName.Password = password;
            port.ClientCredentials.ServiceCertificate.SslCertificateAuthentication = new X509ServiceCertificateAuthentication();

            port.ClientCredentials.ServiceCertificate.SslCertificateAuthentication.CertificateValidationMode = X509CertificateValidationMode.None;
            //port.ClientCredentials.ServiceCertificate.SslCertificateAuthentication.CertificateValidationMode = X509CertificateValidationMode.PeerOrChainTrust;
            localization = new Localization();
            localization.Locale =  CultureInfo.CurrentCulture.Name;
            localization.Timezone = GetTimezone();

            return port;
        }

        public static Localization GetLocalization()
        {
            return localization;
        }

        private static string GetTimezone()
        {
            System.TimeZone tz = TimeZone.CurrentTimeZone;
            System.TimeSpan tspan = tz.GetUtcOffset(System.DateTime.Now);

            // TimeZone.  Format should be '+|-HH:MM' (e.g., -07:00).
            String tzformat;
            if (tspan.Hours >= 0)
            {
                tzformat = String.Format("+{0}:{1}", tspan.Hours.ToString("D2"), tspan.Minutes.ToString("D2"));
            }
            else
            {
                tzformat = String.Format("{0}:{1}", tspan.Hours.ToString("D2"), tspan.Minutes.ToString("D2"));
            }
            return tzformat;
        }
    }
}
