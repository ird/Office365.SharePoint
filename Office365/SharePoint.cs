/*
 * Office365.SharePoint
 *
 *  Created on: 01/02/2014
 *      Author: Stacey Richards
 *
 * The author disclaims copyright to this source code. In place of a legal notice, here is a blessing:
 *
 *    May you do good and not evil.
 *    May you find forgiveness for yourself and forgive others.
 *    May you share freely, never taking more than you give.
 */

using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using System.Xml;

namespace Office365
{
    public class SharePoint
    {
        private delegate void RequestCallback(HttpWebRequest httpWebRequest, object o);

        private delegate string ResponseCallback(Stream stream, object o);

        private string Host;

        private string Site;

        private string Username;

        private CookieContainer Cookies;

        private string RequestDigest;

        private int Expiry;

        private string FileUrl(string serverRelativeUrl)
        {
            int i = serverRelativeUrl.LastIndexOf("/");
            string remoteName;
            if (i == -1)
            {
                remoteName = serverRelativeUrl;
                serverRelativeUrl = "";
            }
            else
            {
                remoteName = serverRelativeUrl.Substring(i + 1);
                serverRelativeUrl = serverRelativeUrl.Substring(0, i);
            }
            return Host + Site + "_api/web/getfolderbyserverrelativeurl('" + serverRelativeUrl + "')/files('" + remoteName + "')";
        }

        private string FolderUrl(string serverRelativeUrl)
        {
            return Host + Site + "_api/web/getfolderbyserverrelativeurl('" + serverRelativeUrl + "')";
        }

        public string CreateFolder(string serverRelativeUrl)
        {
            return Request(Host + Site + "_api/web/folders", null, false, null, true, "POST", "{'__metadata': { 'type': 'SP.Folder' }, 'ServerRelativeUrl': '" + serverRelativeUrl + "'}", "application/json;odata=verbose", null, null, null);
        }

        public string DeleteFile(string serverRelativeUrl)
        {
            return Request(FileUrl(serverRelativeUrl), null, false, "DELETE", true, "POST", null, null, null, null, null);
        }

        public string DeleteFolder(string serverRelativeUrl)
        {
            int i = serverRelativeUrl.LastIndexOf("/");
            string remoteName;
            if (i == -1)
            {
                remoteName = serverRelativeUrl;
                serverRelativeUrl = "";
            }
            else
            {
                remoteName = serverRelativeUrl.Substring(i + 1);
                serverRelativeUrl = serverRelativeUrl.Substring(0, i);
            }
            return Request(FolderUrl(serverRelativeUrl) + "/folders('" + remoteName + "')", null, false, "DELETE", true, "POST", null, null, null, null, null);
        }

        public string DownloadFile(string serverRelativeUrl, string localName)
        {
            return Request(FileUrl(serverRelativeUrl) + "/$value", null, true, null, false, "GET", null, null, null, DownloadFileResponseCallback, localName);
        }

        private string DownloadFileResponseCallback(Stream stream, object o)
        {
            string localName = (string)o;
            byte[] buffer = new byte[1024 * 1024];
            using (FileStream fileStream = new FileStream(localName, FileMode.Create))
            {
                while (true)
                {
                    int bytesRead = stream.Read(buffer, 0, 1024 * 1024);
                    if (bytesRead == 0)
                    {
                        break;
                    }
                    fileStream.Write(buffer, 0, bytesRead);
                }
                fileStream.Close();
            }
            return "";
        }

        public string GetFile(string serverRelativeUrl)
        {
            return Request(FileUrl(serverRelativeUrl), null, false, null, false, "GET", null, null, null, null, null);
        }

        public string GetFileProperty(string serverRelativeUrl, string propertyName)
        {
             return Request(FileUrl(serverRelativeUrl) + "/" + propertyName, null, false, null, false, "GET", null, null, null, null, null);
        }

        public string GetFiles(string serverRelativeUrl)
        {
            return Request(FolderUrl(serverRelativeUrl) + "/files", null, false, null, false, "GET", null, null, null, null, null);
        }

        public string GetFolders(string serverRelativeUrl)
        {
            return Request(FolderUrl(serverRelativeUrl) + "/folders", null, false, null, false, "GET", null, null, null, null, null);
        }

        private String GetRequestDigest()
        {
            Int32 secondsSinceTheUnixEpoch = (Int32)(DateTime.UtcNow - new DateTime(1970, 1, 1)).TotalSeconds;
            if (secondsSinceTheUnixEpoch + 60 > Expiry)
            {
                String s = Request(Host + Site + "_api/contextinfo", null, false, null, false, "POST", null, null, null, null, null);
                XmlReader xmlReader = XmlReader.Create(new MemoryStream(Encoding.UTF8.GetBytes(s)));
                xmlReader.ReadToFollowing("d:FormDigestTimeoutSeconds");
                Expiry = secondsSinceTheUnixEpoch + xmlReader.ReadElementContentAsInt();
                /*
                 * The element following d:FormDigestTimeoutSeconds is d:FormDigestValue. Now that the contents of d:FormDigestTimeoutSeconds
                 * have been read, the xmlReader will be pointing to d:FormDigestValue. Because of this, we don't need to 
                 * xmlReader.ReadToFollowing("d:FormDigestValue"). As a matter of fact, this would be an error because it would be asking the
                 * xmlReader to read the next matching d:formDigestValue tag AFTER the element that it's currently on. The element that
                 * the xmlReader is currently on is the one we want.
                 */
                RequestDigest = xmlReader.ReadElementContentAsString();
            }
            return RequestDigest;
        }

        private string Request(string uri, string accept, bool acceptCompression, string xHttpMethod, bool xRequestDigest, string method, string content, string contentType, RequestCallback requestCallback, ResponseCallback responseCallback, object o)
        {
            HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create(uri);
            if (accept != null && accept != "")
            {
                httpWebRequest.Accept = accept;
            }
            httpWebRequest.AllowAutoRedirect = false;
            if (Cookies != null)
            {
                httpWebRequest.CookieContainer = Cookies;
            }
            if (acceptCompression)
            {
                httpWebRequest.Headers.Add("Accept-Encoding", "gzip, deflate");
            }
            if (xHttpMethod != null && xHttpMethod != "")
            {
                httpWebRequest.Headers.Add("X-HTTP-Method", xHttpMethod);
            }
            if (xRequestDigest)
            {
                httpWebRequest.Headers.Add("X-RequestDigest", GetRequestDigest());
            }
            httpWebRequest.KeepAlive = false;
            if (method != null && method != "")
            {
                httpWebRequest.Method = method;
            }
            if (method == "POST")
            {
                if (content != null && content != "")
                {
                    byte[] bytes = Encoding.UTF8.GetBytes(content);
                    httpWebRequest.ContentLength = content.Length;
                    if (contentType != null && contentType != "")
                    {
                        httpWebRequest.ContentType = contentType;
                    }
                    httpWebRequest.GetRequestStream().Write(bytes, 0, bytes.Length);
                }
                else if (requestCallback != null)
                {
                    requestCallback(httpWebRequest, o);
                }
                else
                {
                    httpWebRequest.ContentLength = 0;
                }
            }
            string s;
            using (WebResponse WebResponse = httpWebRequest.GetResponse())
            {
                using (Stream stream = WebResponse.GetResponseStream())
                {
                    if (responseCallback != null)
                    {
                        s = responseCallback(stream, o);
                    }
                    else
                    {
                        s = new StreamReader(stream).ReadToEnd();
                    }
                }
            }
            return s;
        }

        public List<RoleDefinition> RoleDefinitions()
        {
            List<RoleDefinition> roleDefinitions = new List<RoleDefinition>();
            string s = Request(Host + Site + "_api/web/roledefinitions", null, false, null, false, null, null, null, null, null, null);
            XmlReader xmlReader = XmlReader.Create(new MemoryStream(Encoding.UTF8.GetBytes(s)));
            while (xmlReader.ReadToFollowing("m:properties"))
            {
                xmlReader.ReadToDescendant("d:High");
                RoleDefinition roleDefinition = new RoleDefinition();
                roleDefinition.BasePermissions = new BasePermissions();
                roleDefinition.BasePermissions.High = xmlReader.ReadElementContentAsLong();
                roleDefinition.BasePermissions.Low = xmlReader.ReadElementContentAsLong();
                xmlReader.ReadToFollowing("d:Description");
                roleDefinition.Description = xmlReader.ReadElementContentAsString();
                roleDefinition.Hidden = xmlReader.ReadElementContentAsBoolean();
                roleDefinition.Id = xmlReader.ReadElementContentAsInt();
                roleDefinition.Name = xmlReader.ReadElementContentAsString();
                roleDefinition.Order = xmlReader.ReadElementContentAsInt();
                roleDefinition.RoleTypeKind = xmlReader.ReadElementContentAsInt();
                roleDefinitions.Add(roleDefinition);
            }
            return roleDefinitions;
        }

        /// <summary>
        /// http://macfoo.wordpress.com/
        /// http://allthatjs.com/2012/03/28/remote-authentication-in-sharepoint-online/
        /// </summary>
        /// <param name="host"></param>
        /// <param name="site"></param>
        /// <param name="username"></param>
        /// <param name="password"></param>
        public SharePoint(string host, string site, string username, string password)
        {
            Host = host;
            Site = site;
            Username = username;
            Cookies = null;
            RequestDigest = null;
            Expiry = 0;
            string s = Request(
                "https://login.microsoftonline.com/extSTS.srf",
                null,
                false,
                null,
                false,
                "POST",
                string.Format(
                    "<s:Envelope xmlns:s=\"http://www.w3.org/2003/05/soap-envelope\" xmlns:a=\"http://www.w3.org/2005/08/addressing\" xmlns:u=\"http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd\">\r\n" +
                        "<s:Header>\r\n" +
                        "<a:Action s:mustUnderstand=\"1\">http://schemas.xmlsoap.org/ws/2005/02/trust/RST/Issue</a:Action>\r\n" +
                        "<a:ReplyTo>\r\n" +
                            "<a:Address>http://www.w3.org/2005/08/addressing/anonymous</a:Address>\r\n" +
                        "</a:ReplyTo>\r\n" +
                        "<a:To s:mustUnderstand=\"1\">https://login.microsoftonline.com/extSTS.srf</a:To>\r\n" +
                        "<o:Security s:mustUnderstand=\"1\"\r\n" +
                            "xmlns:o=\"http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd\">\r\n" +
                            "<o:UsernameToken>\r\n" +
                            "<o:Username>{0}</o:Username>\r\n" +
                            "<o:Password>{1}</o:Password>\r\n" +
                            "</o:UsernameToken>\r\n" +
                        "</o:Security>\r\n" +
                        "</s:Header>\r\n" +
                        "<s:Body>\r\n" +
                        "<t:RequestSecurityToken xmlns:t=\"http://schemas.xmlsoap.org/ws/2005/02/trust\">\r\n" +
                            "<wsp:AppliesTo xmlns:wsp=\"http://schemas.xmlsoap.org/ws/2004/09/policy\">\r\n" +
                            "<a:EndpointReference>\r\n" +
                                "<a:Address>{2}</a:Address>\r\n" +
                            "</a:EndpointReference>\r\n" +
                            "</wsp:AppliesTo>\r\n" +
                            "<t:KeyType>http://schemas.xmlsoap.org/ws/2005/05/identity/NoProofKey</t:KeyType>\r\n" +
                            "<t:RequestType>http://schemas.xmlsoap.org/ws/2005/02/trust/Issue</t:RequestType>\r\n" +
                            "<t:TokenType>urn:oasis:names:tc:SAML:1.0:assertion</t:TokenType>\r\n" +
                        "</t:RequestSecurityToken>\r\n" +
                        "</s:Body>\r\n" +
                    "</s:Envelope>\r\n",
                    Username,
                    password,
                    Host),
                null,
                null,
                null,
                null);
            XmlReader xmlReader = XmlReader.Create(new MemoryStream(Encoding.UTF8.GetBytes(s)));
            xmlReader.ReadToFollowing("wsse:BinarySecurityToken");
            Cookies = new CookieContainer();
            Request(Host + "/_forms/default.aspx?wa=wsignin1.0", null, false, null, false, "POST", xmlReader.ReadElementContentAsString(), null, null, null, null);
        }

        public List<User> SiteUsers()
        {
            List<User> users = new List<User>();
            string s = Request(Host + Site + "_api/web/siteusers", null, false, null, false, "GET", null, null, null, null, null);
            XmlReader xmlReader = XmlReader.Create(new MemoryStream(Encoding.UTF8.GetBytes(s)));
            while (xmlReader.ReadToFollowing("m:properties"))
            {
                xmlReader.ReadToDescendant("d:Id");
                User user = new User();
                user.Id = xmlReader.ReadElementContentAsInt();
                user.IsHiddenInUi = xmlReader.ReadElementContentAsBoolean();
                user.LoginName = xmlReader.ReadElementContentAsString();
                user.Title = xmlReader.ReadElementContentAsString();
                user.PrincipalType = xmlReader.ReadElementContentAsInt();
                user.Email = xmlReader.ReadElementContentAsString();
                user.IsSiteAdmin = xmlReader.ReadElementContentAsBoolean();
                users.Add(user);
            }
            return users;
        }

        public string UploadFile(string localName, string serverRelativeUrl)
        {
            int i = serverRelativeUrl.LastIndexOf("/");
            string remoteName;
            if (i == -1)
            {
                remoteName = serverRelativeUrl;
                serverRelativeUrl = "";
            }
            else
            {
                remoteName = serverRelativeUrl.Substring(i + 1);
                serverRelativeUrl = serverRelativeUrl.Substring(0, i);
            }
            return Request(FolderUrl(serverRelativeUrl) + "/files/add(url='" + remoteName + "',overwrite=true)", null, false, null, true, "POST", null, null, UploadFileRequestCallback, null, localName);
        }

        private void UploadFileRequestCallback(HttpWebRequest httpWebRequest, object o)
        {
            string localName = (string)o;
            FileInfo fileInfo = new FileInfo(localName);
            httpWebRequest.ContentLength = fileInfo.Length;
            using (Stream stream = httpWebRequest.GetRequestStream())
            {
                byte[] buffer = new byte[1024 * 1024];
                using (FileStream fileStream = new FileStream(localName, FileMode.Open))
                {
                    while (true)
                    {
                        int bytesRead = fileStream.Read(buffer, 0, 1024 * 1024);
                        if (bytesRead == 0)
                        {
                            break;
                        }
                        stream.Write(buffer, 0, bytesRead);
                    }
                    fileStream.Close();
                }
            }
        }
    }
}
