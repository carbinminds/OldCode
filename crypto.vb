{\rtf1\ansi\ansicpg1252\cocoartf1138
{\fonttbl\f0\froman\fcharset0 Times-Roman;}
{\colortbl;\red255\green255\blue255;}
\margl1440\margr1440\vieww10800\viewh8400\viewkind0
\deftab720
\pard\pardeftab720

\f0\fs24 \cf0 Option Explicit On\
Option Strict On\
\
Imports System\
Imports System.Security.Cryptography\
Imports System.Security.Cryptography.X509Certificates\
Imports System.Security.Cryptography.Pkcs\
Imports System.Text \
\
Imports System.Net.ServicePointManager\
\
Public Class Crypto\
\
\'a0\'a0\'a0 Dim ai As AlgorithmIdentifier\
\
\
\
\'a0\'a0\'a0 Public Sub Headers()\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0 Dim msgHeaders As System.Net.ServicePointManager\
\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0 ' only for IIS 7\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0 msgHeaders.CheckCertificateRevocationList = True\
\
\
\'a0\'a0\'a0 End Sub\
\'a0\'a0\'a0 Public Shared Function Sign(ByVal buffer() As Byte, ByVal x509Cert As X509Certificate2) As Byte()\
\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0 Dim content As ContentInfo\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0 Dim signedCMS As SignedCms\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0 Dim cmsSigner As CmsSigner\
\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0 Try\
\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0 ' Setup the data to be signed\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0 content = New ContentInfo(buffer) \
\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0 signedCMS = New SignedCms(content, False)\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0 cmsSigner = New CmsSigner(SubjectIdentifierType.IssuerAndSerialNumber, x509Cert)\
\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0 ' Now create the signature, the signer will sign the data \
\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0 signedCMS.ComputeSignature(cmsSigner)\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0 Return signedCMS.Encode\
\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0 Catch ex As Exception\
\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0 End Try\
\
\'a0\'a0\'a0 End Function\
\
\'a0\'a0\'a0 Private Sub GetCertificate(ByVal certName As X509Certificate2) \
\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0 Dim certStore As X509Store = Nothing\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0 Dim certCollection As X509Certificate2Collection = Nothing\
\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0 Try\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0 ' Instansiate the certificate store from Personal Certificate, and the Current User \
\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0 certStore = New X509Store(StoreName.My, StoreLocation.CurrentUser)\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0 certStore.Open(OpenFlags.ReadOnly)\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0 certCollection = certStore.Certificates.Find(X509FindType.FindBySubjectName , certName, False)\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0 certStore.Close()\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0 If certCollection.Count = 0 Then\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0 Exit Sub\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0 End If\
\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0 Catch ex As CryptographicException\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0 Throw ex \
\'a0\'a0\'a0\'a0\'a0\'a0\'a0 Catch ex1 As Exception\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0 Throw ex1\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0 Finally\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0 certStore.Close()\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0 End Try\
\
\'a0\'a0\'a0 End Sub\
\
\'a0\'a0\'a0 Public Shared Function Encrypt(ByVal buffer() As Byte, ByVal certName As X509Certificate2) As Byte() \
\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0 Dim content As ContentInfo\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0 Dim envelopedCMS As EnvelopedCms\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0 Dim cmsRecipient As CmsRecipient\
\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0 Try\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0 content = New ContentInfo(buffer)\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0 envelopedCMS = New EnvelopedCms(content) \
\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0 ' Pass a certificate to cmsrecipient to Encrypt\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0 cmsRecipient = New CmsRecipient(SubjectIdentifierType.IssuerAndSerialNumber, certName)\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0 envelopedCMS.Encrypt(cmsRecipient) ' Encrypt \
\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0 Return envelopedCMS.Encode ' After the encrypt perform an encode\
\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0 Catch ex As CryptographicException\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0 Throw ex\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0 Catch ex1 As Exception\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0 Throw ex1\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0 End Try\
\
\'a0\'a0\'a0 End Function\
\
\'a0\'a0\'a0 Public Shared Function Decrypt(ByVal buffer() As Byte) As Byte()\
\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0 Dim envelopedCMS As EnvelopedCms\
\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0 Try\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0 envelopedCMS = New EnvelopedCms() \
\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0 envelopedCMS.Decode(buffer) ' Convert the PKCS7 encoded data into raw data\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0 envelopedCMS.Decrypt() ' This will decrypt the key embedded using the key identified in message\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0 Return envelopedCMS.ContentInfo.Content ' return the content\
\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0 Catch ex As CryptographicException\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0 Throw ex\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0 Catch ex1 As Exception\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0 Throw ex1\
\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0 End Try\
\'a0\'a0\'a0 End Function \
\
\'a0\'a0\'a0 Public Shared Function Verify(ByVal buffer() As Byte, ByVal certName As X509Certificate2) As Boolean\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0 Dim signedCMS As SignedCms\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0 Dim signTrue As Boolean = False\
\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0 Try\
\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0 signedCMS = New SignedCms() \
\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0 signedCMS.Decode(buffer)\
\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0 signedCMS.CheckSignature(New X509Certificate2Collection(certName), False)\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0 Return True\
\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0 Catch ex As CryptographicException\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0 Return False \
\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0 Throw ex\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0 Catch ex1 As Exception\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0 Throw ex1\
\'a0\'a0\'a0\'a0\'a0\'a0\'a0 End Try\
\'a0\'a0\'a0 End Function\
End Class}